using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Translator;

class Program
{
    static void Main(string[] args)
    {
        Console.ForegroundColor = ConsoleColor.DarkYellow;
        Console.WriteLine("--------------------------------------------------------------------------------------------------------------------");
        Console.WriteLine("IMPORTANT NOTES:"); 
        Console.WriteLine("     1. Ensure that the latest en.json file is located at the path specified in appsettings.json.");
        Console.WriteLine("     2. Verify that the path is correctly set; otherwise, the code will not work.");
        Console.WriteLine("     3. Always begin by generating the Excel file using JSON keys and English values from the en.json file (Step 1).");
        Console.WriteLine("        Additionally, ensure that translations for any newly added JSON keys are included in the Excel.");
        Console.WriteLine("        This guarantees that all new keys have corresponding translations.");
        Console.WriteLine("--------------------------------------------------------------------------------------------------------------------");
        Console.WriteLine();
        Console.ResetColor();

        Console.WriteLine("Select the operation to be performed:");
        Console.WriteLine("1. Read en.json file to generate an Excel with JSON keys and english values.");
        Console.WriteLine("2. Generate JSON data from Excel with translations.");
        Console.WriteLine();
        Console.Write("What would you like to choose?: ");
        var userSelection = Console.ReadLine();
        Console.WriteLine();
        Console.WriteLine("--------------------------------------------------------------------------------------------------------------------");
        Console.WriteLine();

        // Build configuration to read from appsettings.json
        var configuration = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        // Initialize variables for selected language and shorthand
        string selectedLanguage = "";
        string selectedLanguageShortHand = "";

        // Get the en.json file path from the configuration
        string jsonFilePath = configuration["EN_JSON_FilePath"];
        //Get the excel file path from the configuration
        string excelFilePath = configuration["Excel_FilePath"];

        if (userSelection == "1")
        {
            JSONToExcel.ConvertJSONToExcel(configuration, jsonFilePath, excelFilePath);
        }
        else if (userSelection == "2")
        {
            Console.WriteLine("Select appropriate language:");
            Console.WriteLine("1. French");
            Console.WriteLine("2. Spanish");
            Console.WriteLine("3. Dutch");
            Console.WriteLine("4. Polish");
            Console.WriteLine();
            Console.Write("What would you like to choose?: ");
            var selectedLanguageIndex = Console.ReadLine();
            switch (selectedLanguageIndex)
            {
                case "1":
                    selectedLanguage = configuration["French"];
                    selectedLanguageShortHand = configuration["FrenchShorthand"];
                    break;
                case "2":
                    selectedLanguage = configuration["Spanish"];
                    selectedLanguageShortHand = configuration["SpanishShorthand"];
                    break;
                case "3":
                    selectedLanguage = configuration["Dutch"];
                    selectedLanguageShortHand = configuration["DutchShorthand"];
                    break;
                case "4":
                    selectedLanguage = configuration["Polish"];
                    selectedLanguageShortHand = configuration["PolishShorthand"];
                    break;
                default:
                    Console.WriteLine("Invalid selection. Defaulting to French.");
                    selectedLanguage = configuration["French"];
                    selectedLanguageShortHand = configuration["FrenchShorthand"];
                    break;
            }
            string outputFilePath = configuration["Output_Lang_JSON_FolderPath"] + selectedLanguageShortHand + ".json";

            var excelToDataTable = new ExcelToDataTable();
            string sheetName = "JSON Data";

            try
            {
                DataTable dataTable = excelToDataTable.ReadExcelToDataTable(excelFilePath, sheetName);
                JSONUpdater.UpdateJsonWithDataTable(jsonFilePath, dataTable, outputFilePath, "JSON Keys", selectedLanguage);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"Error: {ex.Message}");
                Console.ResetColor();
            }
        }
        else 
        {
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine("Invalid selection. Please select either 1 or 2.");
            Console.ResetColor();
            return;
        }
    }

}
