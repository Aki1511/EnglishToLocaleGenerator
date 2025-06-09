using System;
using System.Data;
using System.IO;
using ClosedXML.Excel;

namespace Translator
{
    class ExcelToDataTable
    {
        public DataTable ReadExcelToDataTable(string filePath, string sheetName)
        {
            // Check if the file exists
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"The file '{filePath}' does not exist.");
            }

            // Create a new DataTable
            var dataTable = new DataTable();

            try
            {
                // Open the Excel workbook
                using var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheet(sheetName);

                if (worksheet == null)
                {
                    throw new ArgumentException($"The sheet '{sheetName}' does not exist in the Excel file.");
                }

                // Read the header row
                var headerRow = worksheet.Row(1);
                int columnCount = headerRow.CellsUsed().Count();
                foreach (var cell in headerRow.CellsUsed())
                {
                    dataTable.Columns.Add(cell.Value.ToString());
                }

                // Read the data rows
                foreach (var row in worksheet.RowsUsed().Skip(1)) // Skip the header row
                {
                    var dataRow = dataTable.NewRow();
                    //int columnIndex = 0;

                    //foreach (var cell in row.CellsUsed())
                    //{
                    //    dataRow[columnIndex] = cell.Value.ToString() != "#N/A" ? cell.Value.ToString() : "";
                    //    columnIndex++;
                    //}
                    for (int columnIndex = 1; columnIndex <= columnCount; columnIndex++)
                    {
                        var cell = row.Cell(columnIndex);
                        var value = cell.Value.ToString();

                        // Handle "#N/A" and empty cells
                        dataRow[columnIndex - 1] = !string.IsNullOrEmpty(value) && value != "#N/A" ? value : "";
                    }

                    dataTable.Rows.Add(dataRow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine($"An error occurred while reading the Excel file: {ex.Message}");
                Console.ResetColor();
                throw;
            }

            return dataTable;
        }
    }
}
