using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace csvtoexcelwithheadersDetails01
{
    class Program
    {
        static void Main(string[] args)
        {
            string flatFilePath = @"D:\CATOVIEW\IN200537_NCA_20240611_143926.txt"; // Path to your flat file
            string excelFilePath = @"D:\CATOVIEW\outputD.xlsx";                    // Path to your output Excel file

            // Define the headers
            string[] headers = { "Record_IDentification", "Detail_Record_Line_Number", "DP_ID", "Client_ID", "Client_Account_Category", "Allotment_Quantity", "Lock_in_Reason_Code", "Lock_in_Release_Date", "Issue_Price", "Issued_Amount", "Paid_up_Price", "Paid_up_Amount", "Filler" }; // Adjust these headers according to your file structure
                                                                                                                                                                                                                                                                                            // Define the width of each field

            // Adjust these headers according to your file structure
                                                                    // Define the width of each field
            int[] fieldWidths = { 2, 7, 8, 8, 2, 18, 2, 8, 18, 18, 18, 18, 12 }; // Adjust these widths according to your file structure

            ConvertFixedWidthFileToExcel(flatFilePath, excelFilePath, headers, fieldWidths);
        }

        static void ConvertFixedWidthFileToExcel(string flatFilePath, string excelFilePath, string[] headers, int[] fieldWidths)
        {
            // Ensure EPPlus can use the non-commercial license
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                // Add headers
                for (int j = 0; j < headers.Length; j++)
                {
                    worksheet.Cells[1, j + 1].Value = headers[j];
                }

                var lines = File.ReadAllLines(flatFilePath);
                for (int i = 1; i < lines.Length; i++) // Start from the third line
                {
                    var line = lines[i];
                    int startIndex = 0;
                    for (int j = 0; j < fieldWidths.Length; j++)
                    {
                        if (startIndex < line.Length)
                        {
                            var field = line.Substring(startIndex, fieldWidths[j]).Trim();
                            worksheet.Cells[i+3, j + 1].Value = field; // Data starts at the third row in Excel
                            startIndex += fieldWidths[j];
                        }
                    }
                }

                package.SaveAs(new FileInfo(excelFilePath));
            }

            Console.WriteLine($"Successfully converted {flatFilePath} to {excelFilePath}");
            System.Diagnostics.Process.Start(@"D:\CATOVIEW\outputD.xlsx");
        }
    }
}