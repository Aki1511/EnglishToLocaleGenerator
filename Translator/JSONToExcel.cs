using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;

namespace Translator
{
    class JSONToExcel
    {
        static public void ConvertJSONToExcel(IConfigurationRoot configuration, string jsonFilePath, string excelFilePath)
        {
            try
            {
                // Read the JSON file content
                string jsonContent = File.ReadAllText(jsonFilePath);

                // Deserialize the JSON content into a JsonDocument
                using var jsonDocument = JsonDocument.Parse(jsonContent);

                // Flatten the JSON into key-value pairs
                var flattenedData = new Dictionary<string, string>();
                FlattenJson(jsonDocument.RootElement, flattenedData);

                // Convert the flattened data to an Excel file
                ConvertToExcel(flattenedData, excelFilePath);

                Console.ForegroundColor = ConsoleColor.DarkGreen;
                Console.WriteLine($"Excel file created successfully at: {excelFilePath}");
                Console.ResetColor();
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Error: The specified file was not found.");
                Console.ResetColor();
            }
            catch (JsonException)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("Error: The file content is not a valid JSON.");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.ResetColor();
            }
        }

        static void FlattenJson(JsonElement element, Dictionary<string, string> result, string prefix = "")
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

        static void ConvertToExcel(Dictionary<string, string> data, string filePath)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("JSON Data");

            // Add headers
            worksheet.Cell(1, 1).Value = "JSON Keys";
            worksheet.Cell(1, 2).Value = "English";
            worksheet.Cell(1, 3).Value = "French";
            worksheet.Cell(1, 4).Value = "Spanish";
            worksheet.Cell(1, 5).Value = "Dutch";
            worksheet.Cell(1, 6).Value = "Polish";

            // Add data
            int row = 2;
            foreach (var kvp in data)
            {
                worksheet.Cell(row, 1).Value = kvp.Key;
                worksheet.Cell(row, 2).Value = kvp.Value;
                row++;
            }

            // Adjust column widths
            worksheet.Columns().AdjustToContents();

            // Save the Excel file
            workbook.SaveAs(filePath);
        }
    }
}
