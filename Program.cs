using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;
using System.Text.Json;
using System.Text;

string pstPath = @"C:\Users\alvar\Githubs\ToDoReader\2024-dec.pst";
string fullJsonPath = @"C:\Users\alvar\Githubs\ToDoReader\todos_full.json";
string simpleJsonPath = @"C:\Users\alvar\Githubs\ToDoReader\todos_simple.json";

try
{
    using (var pst = PersonalStorage.FromFile(pstPath))
    {
        var allTasksFull = new List<Dictionary<string, object>>();
        var allTasksSimple = new List<Dictionary<string, object>>();
        Console.WriteLine($"Processing PST: {pstPath}");
        ProcessFolder(pst, pst.RootFolder, allTasksFull, allTasksSimple);
        
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
        Console.WriteLine($"Full export: {fullJsonPath}");
        Console.WriteLine($"Simple export: {simpleJsonPath}");
    }
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
}

void ProcessFolder(PersonalStorage pst, FolderInfo folder, 
    List<Dictionary<string, object>> allTasksFull, 
    List<Dictionary<string, object>> allTasksSimple)
{
    foreach (var msgInfo in folder.EnumerateMessages())
    {
        if (msgInfo.MessageClass == "IPM.Microsoft.Todo" || 
            msgInfo.MessageClass == "IPM.Task" ||
            msgInfo.MessageClass == "IPM.Todo.Microsoft.Todo.Roamed")
        {
            using (var msg = pst.ExtractMessage(msgInfo))
            {
                // Full task details
                var taskFull = new Dictionary<string, object>
                {
                    ["Subject"] = msg.Subject ?? "",
                    ["MessageClass"] = msgInfo.MessageClass,
                    ["Body"] = msg.Body ?? "",
                    ["CreationTime"] = msg.DeliveryTime,
                    ["LastModifiedTime"] = GetMapiPropertyDateTime(msg, 0x3008) ?? DateTime.MinValue,
                    ["DueDate"] = GetMapiPropertyDateTime(msg, 0x81020040),
                    ["StartDate"] = GetMapiPropertyDateTime(msg, 0x81030040),
                    ["CompletedDate"] = GetMapiPropertyDateTime(msg, 0x810F0040),
                    ["ReminderTime"] = GetMapiPropertyDateTime(msg, 0x8502),
                    ["IsComplete"] = GetMapiPropertyBool(msg, 0x81EC000B) ?? false,
                    ["Categories"] = GetMapiPropertyString(msg, 0x850C001E) ?? "",
                    ["TaskId"] = GetMapiPropertyString(msg, 0x85E0001F) ?? "",
                    ["Priority"] = GetMapiPropertyInt(msg, 0x85150003) ?? 0,
                    ["PercentComplete"] = GetMapiPropertyDouble(msg, 0x81050003) ?? 0.0,
                    ["Status"] = GetMapiPropertyInt(msg, 0x81110003) ?? 0
                };

                // Add all custom properties to full export
                var customProps = new Dictionary<string, object>();
                foreach (var prop in msg.Properties)
                {
                    try
                    {
                        if (!taskFull.ContainsKey(prop.Key.ToString()))
                        {
                            customProps[prop.Key.ToString()] = prop.Value?.GetValue()?.ToString() ?? "";
                        }
                    }
                    catch { }
                }
                taskFull["CustomProperties"] = customProps;
                allTasksFull.Add(taskFull);

                // Simplified task details
                var taskSimple = new Dictionary<string, object>
                {
                    ["Subject"] = msg.Subject ?? "",
                    ["MessageClass"] = msgInfo.MessageClass,
                    ["Body"] = msg.Body ?? "",
                    ["CreationTime"] = msg.DeliveryTime,
                    ["LastModifiedTime"] = GetMapiPropertyDateTime(msg, 0x3008) ?? DateTime.MinValue,
                    ["IsComplete"] = GetMapiPropertyBool(msg, 0x81EC000B) ?? false,
                    ["Categories"] = GetMapiPropertyString(msg, 0x850C001E) ?? "",
                    ["TaskId"] = GetMapiPropertyString(msg, 0x85E0001F) ?? "",
                    ["Status"] = GetMapiPropertyInt(msg, 0x81110003) ?? 0,
                    ["FolderName"] = customProps.GetValueOrDefault("235208735", "")?.ToString() ?? "",
                    ["CreatedTime"] = customProps.GetValueOrDefault("235274304", "")?.ToString() ?? "",
                    ["LocalId"] = customProps.GetValueOrDefault("912392223", "")?.ToString() ?? "",
                    ["Creator"] = customProps.GetValueOrDefault("1073217567", "")?.ToString() ?? ""
                };

                allTasksSimple.Add(taskSimple);

                // Print to console for visibility
                Console.WriteLine($"\n[Task Found]");
                Console.WriteLine($"Title: {taskSimple["Subject"]}");
                Console.WriteLine($"Type: {taskSimple["MessageClass"]}");
            }
        }
    }

    foreach (var subfolder in folder.GetSubFolders())
    {
        ProcessFolder(pst, subfolder, allTasksFull, allTasksSimple);
    }
}

DateTime? GetMapiPropertyDateTime(MapiMessage msg, long tag)
{
    var value = GetMapiProperty<DateTime>(msg, tag);
    return value;
}

bool? GetMapiPropertyBool(MapiMessage msg, long tag)
{
    var value = GetMapiProperty<bool>(msg, tag);
    return value;
}

T? GetMapiProperty<T>(MapiMessage msg, long tag) where T : struct
{
    var prop = msg.Properties[tag];
    return prop != null ? (T?)prop.GetValue() : null;
}

string? GetMapiPropertyString(MapiMessage msg, long tag)
{
    var prop = msg.Properties[tag];
    return prop?.GetString();
}

int? GetMapiPropertyInt(MapiMessage msg, long tag)
{
    var prop = msg.Properties[tag];
    return prop != null ? (int?)prop.GetValue() : null;
}

double? GetMapiPropertyDouble(MapiMessage msg, long tag)
{
    var prop = msg.Properties[tag];
    return prop != null ? (double?)prop.GetValue() : null;
}