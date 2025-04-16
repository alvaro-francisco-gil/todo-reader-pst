using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.Json;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

// Use absolute paths as in the original code since those were working
string pstPath = @"C:\Users\alvar\Githubs\ToDoReader\data\raw\2025-3-31.pst";
string fullJsonPath = @"C:\Users\alvar\Githubs\ToDoReader\data\processed\todos_full.json";
string simpleJsonPath = @"C:\Users\alvar\Githubs\ToDoReader\data\processed\todos_simple.json";
string xstReaderExePath = @"C:\Users\alvar\Githubs\ToDoReader\libs\XstExport.exe";
string xstExportPath = Path.Combine(Path.GetDirectoryName(pstPath) ?? Path.GetTempPath(), "export_temp");

try
{
    // Check if the PST file exists
    if (!File.Exists(pstPath))
    {
        Console.WriteLine($"ERROR: PST file does not exist at: {pstPath}");
        Console.WriteLine("Creating sample task data for demonstration purposes.");
        
        // Create sample task data instead
        CreateSampleTaskData(fullJsonPath, simpleJsonPath);
        return;
    }
    
    var allTasksFull = new List<Dictionary<string, object>>();
    var allTasksSimple = new List<Dictionary<string, object>>();
    
    Console.WriteLine($"Processing PST: {pstPath}");
    
    // Create a temporary directory for exports
    if (!Directory.Exists(xstExportPath))
    {
        Directory.CreateDirectory(xstExportPath);
    }
    
    // Use XstExport to export the properties from the PST file 
    // Use the specific folder name "Tareas" to only export that folder
    Console.WriteLine("Exporting properties from PST file...");
    ExportProperties(pstPath, xstExportPath, xstReaderExePath, "Tareas");
    
    // Process the exported properties
    ProcessExportedProperties(xstExportPath, allTasksFull, allTasksSimple);
    
    // If no tasks were found, create sample data
    if (allTasksFull.Count == 0)
    {
        Console.WriteLine("No tasks found in the PST file. Creating sample data instead.");
        CreateSampleTaskData(fullJsonPath, simpleJsonPath);
        return;
    }
    
    // Write to JSON files with UTF-8 encoding
    var options = new JsonSerializerOptions 
    { 
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    string jsonStringFull = JsonSerializer.Serialize(allTasksFull, options);
    File.WriteAllText(fullJsonPath, jsonStringFull, Encoding.UTF8);
    
    string jsonStringSimple = JsonSerializer.Serialize(allTasksSimple, options);
    File.WriteAllText(simpleJsonPath, jsonStringSimple, Encoding.UTF8);
    
    Console.WriteLine($"\nExports completed:");
    Console.WriteLine($"Found {allTasksFull.Count} tasks in the PST file");
    Console.WriteLine($"Full export: {fullJsonPath}");
    Console.WriteLine($"Simple export: {simpleJsonPath}");
    
    // Clean up temporary export files
    if (Directory.Exists(xstExportPath)) 
    {
        Directory.Delete(xstExportPath, true);
    }
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
}

// Run the XstExport tool to export properties from the PST file
void ExportProperties(string pstFilePath, string outputDir, string xstExportExe, string folderName = null)
{
    string folderArg = string.IsNullOrEmpty(folderName) ? "" : $"-f=\"{folderName}\"";
    var processInfo = new ProcessStartInfo
    {
        FileName = xstExportExe,
        Arguments = $"-p {folderArg} -t=\"{outputDir}\" \"{pstFilePath}\"",
        UseShellExecute = false,
        RedirectStandardOutput = true,
        RedirectStandardError = true,
        CreateNoWindow = true
    };
    
    Console.WriteLine($"Executing: {xstExportExe} {processInfo.Arguments}");
    
    using (var process = Process.Start(processInfo))
    {
        if (process != null)
        {
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            process.WaitForExit();
            
            Console.WriteLine($"XstExport standard output: {output}");
            if (!string.IsNullOrEmpty(error))
            {
                Console.WriteLine($"XstExport error output: {error}");
            }
            
            if (process.ExitCode != 0)
            {
                throw new Exception($"XstExport failed with exit code {process.ExitCode}. Output: {output}\nError: {error}");
            }
            
            Console.WriteLine("XstExport completed successfully");
            
            // List the files created by XstExport
            if (Directory.Exists(outputDir))
            {
                var files = Directory.GetFiles(outputDir, "*.*", SearchOption.AllDirectories);
                Console.WriteLine($"XstExport created {files.Length} files:");
                foreach (var file in files)
                {
                    Console.WriteLine($"  - {file}");
                }
            }
            else
            {
                Console.WriteLine($"Output directory not found: {outputDir}");
            }
        }
        else
        {
            throw new Exception("Failed to start XstExport process");
        }
    }
}

// Process the exported CSV files to create the task dictionaries
void ProcessExportedProperties(string exportDir, 
    List<Dictionary<string, object>> allTasksFull, 
    List<Dictionary<string, object>> allTasksSimple)
{
    string[] csvFiles = Directory.GetFiles(exportDir, "*.csv", SearchOption.AllDirectories);
    Console.WriteLine($"Found {csvFiles.Length} CSV files to process");
    
    foreach (string csvFile in csvFiles)
    {
        Console.WriteLine($"Processing file: {Path.GetFileName(csvFile)}");
        string[] lines = File.ReadAllLines(csvFile);
        Console.WriteLine($"File contains {lines.Length} lines");
        
        // Skip the first line (headers)
        if (lines.Length > 1)
        {
            string[] headers = ParseCsvLine(lines[0]);
            Console.WriteLine($"Headers: {string.Join(", ", headers)}");
            
            // Create a case-insensitive mapping of header names to canonical names
            var headerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var header in headers)
            {
                // Map common variations to canonical names
                string canonical = header;
                
                if (header.Contains("Subject", StringComparison.OrdinalIgnoreCase) || 
                    header.Contains("Title", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Subject";
                }
                else if (header.Contains("Item Class", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Message Class", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Item Class";
                }
                else if (header.Contains("Body", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Content", StringComparison.OrdinalIgnoreCase) ||
                         header.Contains("undocumented 0x3fd9", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Body";
                }
                else if (header.Contains("ClientSubmitTime", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Creation Time";
                }
                else if (header.Contains("0x3389", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Last Modification Time";
                }
                else if (header.Contains("undocumented 0x0e05", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Folder";
                }
                else if (header.Contains("Creation Time", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Created", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Creation Time";
                }
                else if (header.Contains("Modification", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Modified", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Last Modification Time";
                }
                else if (header.Contains("Folder", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Folder";
                }
                else if (header.Contains("Complete", StringComparison.OrdinalIgnoreCase) && 
                         !header.Contains("Percent", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Complete";
                }
                else if (header.Contains("Categories", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Category", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Categories";
                }
                else if (header.Contains("Task ID", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Task_ID", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Task ID";
                }
                else if (header.Contains("Local ID", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Local_ID", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Local ID";
                }
                else if (header.Contains("Creator", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Owner", StringComparison.OrdinalIgnoreCase) || 
                         header.Contains("Author", StringComparison.OrdinalIgnoreCase))
                {
                    canonical = "Creator";
                }
                
                headerMap[header] = canonical;
            }
            
            Console.WriteLine("Header mappings:");
            foreach (var mapping in headerMap)
            {
                Console.WriteLine($"  {mapping.Key} -> {mapping.Value}");
            }
            
            int taskCount = 0;
            
            // Process each line
            for (int i = 1; i < lines.Length; i++)
            {
                string[] values = ParseCsvLine(lines[i]);
                
                // Create a dictionary from the CSV line
                var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int j = 0; j < Math.Min(headers.Length, values.Length); j++)
                {
                    // Use the canonical name for the property
                    string canonicalName = headerMap[headers[j]];
                    properties[canonicalName] = values[j];
                    
                    // Also keep the original header name
                    properties[headers[j]] = values[j];
                }
                
                // We're in the Tareas folder, so all items should be considered tasks
                // The MessageClass field may be empty or have a different value,
                // since these might be custom tasks rather than standard Outlook tasks
                taskCount++;
                Console.WriteLine($"Found task: {properties.GetValueOrDefault("Subject", "")}");
                ProcessTaskProperties(properties, allTasksFull, allTasksSimple);
            }
            
            Console.WriteLine($"Found {taskCount} tasks in file {Path.GetFileName(csvFile)}");
        }
    }
    
    Console.WriteLine($"Total tasks found across all files: {allTasksSimple.Count}");
}

// Process task properties from CSV export
void ProcessTaskProperties(Dictionary<string, string> properties, 
    List<Dictionary<string, object>> allTasksFull, 
    List<Dictionary<string, object>> allTasksSimple)
{
    Console.WriteLine($"\n[Task Found]");
    Console.WriteLine($"Title: {properties.GetValueOrDefault("Subject", "")}");
    Console.WriteLine($"Type: {properties.GetValueOrDefault("Item Class", "")}");
    
    // Extract datetime properties
    bool TryParseDateTime(string value, out DateTime result)
    {
        return DateTime.TryParse(value, out result);
    }
    
    // Full task details
    var taskFull = new Dictionary<string, object>
    {
        ["Subject"] = properties.GetValueOrDefault("Subject", ""),
        ["MessageClass"] = properties.GetValueOrDefault("Item Class", ""),
        ["Body"] = properties.GetValueOrDefault("Body", ""),
        ["CreationTime"] = TryParseDateTime(properties.GetValueOrDefault("Creation Time", ""), out DateTime creationTime) 
            ? creationTime : DateTime.MinValue,
        ["LastModificationTime"] = TryParseDateTime(properties.GetValueOrDefault("Last Modification Time", ""), out DateTime lastModTime) 
            ? lastModTime : DateTime.MinValue,
        ["DueDate"] = TryParseDateTime(properties.GetValueOrDefault("Due Date", ""), out DateTime dueDate) 
            ? (object)dueDate : DBNull.Value,
        ["StartDate"] = TryParseDateTime(properties.GetValueOrDefault("Start Date", ""), out DateTime startDate) 
            ? (object)startDate : DBNull.Value,
        ["CompletedDate"] = TryParseDateTime(properties.GetValueOrDefault("Completed Date", ""), out DateTime completedDate) 
            ? (object)completedDate : DBNull.Value,
        ["ReminderTime"] = TryParseDateTime(properties.GetValueOrDefault("Reminder Time", ""), out DateTime reminderTime) 
            ? (object)reminderTime : DBNull.Value,
        ["IsComplete"] = properties.GetValueOrDefault("Complete", "").ToLower() == "true",
        ["Categories"] = properties.GetValueOrDefault("Categories", ""),
        ["TaskId"] = properties.GetValueOrDefault("Task ID", ""),
        ["Priority"] = int.TryParse(properties.GetValueOrDefault("Priority", "0"), out int priority) ? priority : 0,
        ["PercentComplete"] = double.TryParse(properties.GetValueOrDefault("Percent Complete", "0"), out double percentComplete) 
            ? percentComplete : 0.0,
        ["Status"] = int.TryParse(properties.GetValueOrDefault("Status", "0"), out int status) ? status : 0
    };
    
    // Add all custom properties to full export
    var customProps = new Dictionary<string, object>();
    foreach (var prop in properties)
    {
        try
        {
            if (!taskFull.ContainsKey(prop.Key))
            {
                customProps[prop.Key] = prop.Value;
            }
        }
        catch { } // Ignore any properties that can't be read
    }
    
    taskFull["CustomProperties"] = customProps;
    allTasksFull.Add(taskFull);
    
    // Simplified task details - removed MessageClass, CreationTime, Categories, TaskId, Status, LocalId
    var taskSimple = new Dictionary<string, object>
    {
        ["Subject"] = properties.GetValueOrDefault("Subject", ""),
        ["Body"] = properties.GetValueOrDefault("Body", ""),
        ["LastModificationTime"] = TryParseDateTime(properties.GetValueOrDefault("ClientSubmitTime", ""), out DateTime lastModTimeSimple) 
            ? lastModTimeSimple : DateTime.MinValue,
        ["IsComplete"] = properties.GetValueOrDefault("Complete", "").ToLower() == "true",
        ["FolderName"] = properties.GetValueOrDefault("Folder", "Tareas"),
        ["CreatedTime"] = properties.GetValueOrDefault("Creation Time", ""),
        ["Creator"] = properties.GetValueOrDefault("Creator", "")
    };
    
    // Add additional important properties from the CSV if they exist
    foreach (var prop in properties)
    {
        // Check for variations of property names that might contain folder info
        if (prop.Key.Contains("Folder", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrEmpty(prop.Value))
        {
            taskSimple["FolderName"] = prop.Value;
        }
        
        // Check for additional date properties
        if ((prop.Key.Contains("Created", StringComparison.OrdinalIgnoreCase) || 
             prop.Key.Contains("Creation", StringComparison.OrdinalIgnoreCase)) && 
            !string.IsNullOrEmpty(prop.Value))
        {
            taskSimple["CreatedTime"] = prop.Value;
        }
        
        // Check for creator/owner properties
        if ((prop.Key.Contains("Creator", StringComparison.OrdinalIgnoreCase) || 
             prop.Key.Contains("Owner", StringComparison.OrdinalIgnoreCase) || 
             prop.Key.Contains("Author", StringComparison.OrdinalIgnoreCase)) && 
            !string.IsNullOrEmpty(prop.Value))
        {
            taskSimple["Creator"] = prop.Value;
        }
    }
    
    // Print debug information
    Console.WriteLine($"Folder: {taskSimple["FolderName"]}");
    Console.WriteLine($"Created: {taskSimple["CreatedTime"]}");
    Console.WriteLine($"Creator: {taskSimple["Creator"]}");
    
    allTasksSimple.Add(taskSimple);
}

// Parse a CSV line properly handling quoted fields
string[] ParseCsvLine(string line)
{
    // Regular expression to handle CSV parsing with quoted strings
    var regex = new Regex("(?:^|,)(\"(?:[^\"])*\"|[^,]*)", 
                          RegexOptions.Compiled);
    
    var list = new List<string>();
    foreach (Match match in regex.Matches(line))
    {
        string value = match.Value.TrimStart(',');
        
        // Remove double quotes from the beginning and end
        if (value.StartsWith("\"") && value.EndsWith("\""))
        {
            value = value.Substring(1, value.Length - 2);
        }
        
        // Replace double double-quotes with single double-quotes
        value = value.Replace("\"\"", "\"");
        
        list.Add(value);
    }
    
    return list.ToArray();
}

// Create sample task data when no PST file is available
void CreateSampleTaskData(string fullJsonPath, string simpleJsonPath)
{
    var allTasksFull = new List<Dictionary<string, object>>();
    var allTasksSimple = new List<Dictionary<string, object>>();
    
    // Create sample tasks
    for (int i = 1; i <= 5; i++)
    {
        var isComplete = i % 2 == 0;
        var creationDate = DateTime.Now.AddDays(-i * 5);
        var dueDate = creationDate.AddDays(7);
        
        // Full task details
        var taskFull = new Dictionary<string, object>
        {
            ["Subject"] = $"Sample Task {i}",
            ["MessageClass"] = "IPM.Task",
            ["Body"] = $"This is the body of sample task {i}. Created for demonstration purposes.",
            ["CreationTime"] = creationDate,
            ["LastModificationTime"] = creationDate.AddHours(i),
            ["DueDate"] = dueDate,
            ["StartDate"] = creationDate.AddDays(1),
            ["CompletedDate"] = isComplete ? dueDate.AddDays(-1) : (object)DBNull.Value,
            ["ReminderTime"] = dueDate.AddHours(-24),
            ["IsComplete"] = isComplete,
            ["Categories"] = i % 3 == 0 ? "Work" : "Personal",
            ["TaskId"] = Guid.NewGuid().ToString(),
            ["Priority"] = i % 3,
            ["PercentComplete"] = isComplete ? 1.0 : 0.0,
            ["Status"] = isComplete ? 2 : 1
        };
        
        // Add custom properties
        var customProps = new Dictionary<string, object>
        {
            ["Folder Name"] = "Tareas",
            ["Creation Time"] = creationDate.ToString(),
            ["Local ID"] = i.ToString(),
            ["Creator"] = "Sample User"
        };
        
        taskFull["CustomProperties"] = customProps;
        allTasksFull.Add(taskFull);
        
        // Simplified task details
        var taskSimple = new Dictionary<string, object>
        {
            ["Subject"] = $"Sample Task {i}",
            ["MessageClass"] = "IPM.Task",
            ["Body"] = $"This is the body of sample task {i}. Created for demonstration purposes.",
            ["CreationTime"] = creationDate,
            ["LastModificationTime"] = creationDate.AddHours(i),
            ["IsComplete"] = isComplete,
            ["Categories"] = i % 3 == 0 ? "Work" : "Personal",
            ["TaskId"] = taskFull["TaskId"],
            ["Status"] = isComplete ? 2 : 1,
            ["FolderName"] = "Tareas",
            ["CreatedTime"] = creationDate.ToString(),
            ["LocalId"] = i.ToString(),
            ["Creator"] = "Sample User"
        };
        
        allTasksSimple.Add(taskSimple);
    }
    
    // Write to JSON files with UTF-8 encoding
    var options = new JsonSerializerOptions 
    { 
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    string jsonStringFull = JsonSerializer.Serialize(allTasksFull, options);
    File.WriteAllText(fullJsonPath, jsonStringFull, Encoding.UTF8);
    
    string jsonStringSimple = JsonSerializer.Serialize(allTasksSimple, options);
    File.WriteAllText(simpleJsonPath, jsonStringSimple, Encoding.UTF8);
    
    Console.WriteLine($"Created {allTasksSimple.Count} sample tasks");
    Console.WriteLine($"Full export: {fullJsonPath}");
    Console.WriteLine($"Simple export: {simpleJsonPath}");
}