using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using Microsoft.Extensions.Configuration;

namespace Translator
{
    class JSONUpdater
    {
        public static void UpdateJsonWithDataTable(string jsonFilePath, DataTable dataTable, string outputFilePath, string keyColumnName, string valueColumnName)
        {
            try
            {
                // Read the JSON file content
                string jsonContent = File.ReadAllText(jsonFilePath);

                // Parse the JSON into a JsonDocument
                using var jsonDocument = JsonDocument.Parse(jsonContent);
                var jsonElement = jsonDocument.RootElement;

                // Convert the JSON to a mutable dictionary
                var jsonDictionary = new Dictionary<string, object>();
                FlattenJson(jsonElement, jsonDictionary);

                // Update the JSON dictionary with values from the DataTable
                foreach (DataRow row in dataTable.Rows)
                {
                    string key = row[keyColumnName].ToString(); // Replace "keyColumnName" with the actual column name containing the JSON keys
                    string value = row[valueColumnName].ToString(); // Replace "valueColumnName" with the actual column name containing the replacement values

                    if (jsonDictionary.ContainsKey(key) && !String.IsNullOrEmpty(value))
                    {
                        jsonDictionary[key] = value;
                    }
                }

                // Rebuild the JSON object from the updated dictionary
                var updatedJson = RebuildJson(jsonDictionary);

                // Write the updated JSON to the output file with special character retention
                var serializerOptions = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                File.WriteAllText(outputFilePath, JsonSerializer.Serialize(updatedJson, serializerOptions));
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"JSON file updated successfully at: {outputFilePath}");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"An error occurred while updating the JSON file: {ex.Message}");
                Console.ResetColor();
            }
        }

        private static void FlattenJson(JsonElement element, Dictionary<string, object> result, string prefix = "")
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    foreach (var property in element.EnumerateObject())
                    {
                        string newPrefix = string.IsNullOrEmpty(prefix) ? property.Name : $"{prefix}.{property.Name}";
                        FlattenJson(property.Value, result, newPrefix);
                    }
                    break;

                case JsonValueKind.Array:
                    int index = 0;
                    foreach (var item in element.EnumerateArray())
                    {
                        FlattenJson(item, result, $"{prefix}[{index}]");
                        index++;
                    }
                    break;

                default:
                    result[prefix] = element.ToString();
                    break;
            }
        }

        private static Dictionary<string, object> RebuildJson(Dictionary<string, object> flattenedJson)
        {
            var result = new Dictionary<string, object>();

            foreach (var kvp in flattenedJson)
            {
                var keys = kvp.Key.Split('.');
                var current = result;

                for (int i = 0; i < keys.Length; i++)
                {
                    if (i == keys.Length - 1)
                    {
                        current[keys[i]] = kvp.Value;
                    }
                    else
                    {
                        if (!current.ContainsKey(keys[i]))
                        {
                            current[keys[i]] = new Dictionary<string, object>();
                        }

                        current = (Dictionary<string, object>)current[keys[i]];
                    }
                }
            }

            return result;
        }
    }
}
