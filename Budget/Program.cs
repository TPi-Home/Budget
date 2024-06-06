using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace BudgetSpreadsheet
{
    class Program
    {
        static void Main(string[] args)
        {
            string currentYear = DateTime.Now.ToString("yyyy");
            string workbookFileName = $"{currentYear}.xlsx";

            // Initial Menu
            Console.WriteLine("What do you want to do?");
            Console.WriteLine("1. Create a template or add/merge in bills to an existing worksheet.");
            Console.WriteLine("2. Finalize the current sheet and start a new sheet for data entry.");
            Console.WriteLine("3. Open README/TUTORIAL.");
            Console.WriteLine("4. Exit.");
            Console.Write("Enter your choice (1, 2, 3, or 4): ");
            string choice = Console.ReadLine();
            while (choice != "4")
            {
                switch (choice)
                {
                    case "1":
                        // Create a new template workbook
                        // Add bills to an existing workbook
                        // CreateTemplate(workbookFileName);
                        AddBillsToCurrentSheet(workbookFileName);
                        break;
                    case "2":
                        FinalizeCurrentSheet(workbookFileName);
                        break;
                    case "3":
                        ReadMe();
                        break;
                    case "4":
                        Console.WriteLine("Closing.");
                        break;
                    default:
                        Console.WriteLine("Invalid choice.");
                        break;
                }
                Console.WriteLine("What do you want to do?");
                Console.WriteLine("1. Create a template or add/merge in bills to an existing worksheet.");
                Console.WriteLine("2. Finalize the current sheet and start a new sheet for data entry.");
                Console.WriteLine("3. Open README/TUTORIAL.");
                Console.WriteLine("4. Exit.");
                Console.Write("Enter your choice (1, 2, 3, or 4): ");

                //break up rent and mortgage, taxes (plus adjusting applicable bills to tax rate, prob seperate column), seperate insurance types, subscription audit, income and capital gains, 1099 income, savings info, debt payments, cells for tax season reminders
                choice = Console.ReadLine();
            }

        }

        static void ReadMe()
        {
            Console.WriteLine("*This software gathers information from the user and builds a spreadsheet based off that information.\n*This is the style of spreadsheet I have used to track my expenses/budget over the last couple years.\n*While I have been content with this model for my personal use, I do believe there will be more to come as I clean it up some.\n*Long term features in the pipeline: " +
                "\n*More specific information surrounding rent/mortgage\n*Possibly some functionality to chart information relevant to a budget\n*Taxes and figures with taxes factored in\n*More details surrounding subscription costs and how they add up\n*Capital gains or 1099 income, as well as more quality of life for w2 tracking\n*more detailed debt tracking\n*Reminders for tax season (healthcare.gov tax credit, forms" +
                "popping up to remind you to file that part with the rest of your taxes)\n*Please report bugs directly to me if you know me\n \n \n \n Copyright 2024 Tyler Pittman\r\n\r\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\r\n\r\nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\r\n\r\nTHE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. ");
        }
        /*static void CreateTemplate(string workbookFileName) //It seems as if one function is better for both creation and merging because it means I don't have to pass around a dictionary
        {
            using (var workbook = new XLWorkbook())
            {
                string[] args = null;
                // Check for file in root directory
                if (File.Exists(workbookFileName))
                {
                    Console.WriteLine($"You already have a Workbook for this year!");
                    return;
                }
                var worksheet = workbook.Worksheets.Add("Entry");

                // Set up the headers
                worksheet.Cell("A1").Value = "Bill Name";
                worksheet.Cell("B1").Value = "Minimum Amount Owed";
                worksheet.Cell("C1").Value = "Minimum Amount Due";
                worksheet.Cell("D1").Value = "Due Date Week";
                worksheet.Cell("E1").Value = "Transition Formula";
                worksheet.Cell("F1").Value = "Latest Due Date";
                worksheet.Cell("G1").Value = "Autopay Status";
                worksheet.Cell("H1").Value = "Paid Boolean";
                worksheet.Cell("I1").Value = "Amount Paid";

                // Format the headers
                var headerRange = worksheet.Range("A1:I1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                //Column width formatting
                foreach (var column in worksheet.Columns())
                {
                    column.Width = 26;
                }

                List<(string billName, decimal amount, bool isSplit, int week, string autopayStatus)> bills = new List<(string, decimal, bool, int, string)>();
                HashSet<int> addedWeeks = new HashSet<int>();
                for (int week = 1; week <= 4; week++)
                {
                    if (!addedWeeks.Contains(week))
                    {
                        bills.Add(($"Week: {week}", 0, false, week, ""));
                        addedWeeks.Add(week);
                    }
                    Console.Write($"How many bills do you have for week {week}? ");
                    int numberOfBills = int.Parse(Console.ReadLine());

                    for (int i = 0; i < numberOfBills; i++)
                    {
                        Console.Write($"Enter the name of bill {i + 1} for week {week}: ");
                        string billName = Console.ReadLine();
                        Console.Write($"Enter the amount for {billName}: ");
                        decimal amount = decimal.Parse(Console.ReadLine());
                        Console.Write($"Are you splitting {billName} with a roommate? (yes/no): ");
                        bool isSplit = Console.ReadLine().Trim().ToLower() == "yes";
                        Console.Write($"Enter autopay status for {billName} (yes/no): ");
                        string autopayStatus = Console.ReadLine().Trim().ToLower();
                        bills.Add((billName, amount, isSplit, week, autopayStatus));
                    }
                }

                int currentRow = 2; //Row 1 == headers

                foreach (var bill in bills)
                {
                    while (!worksheet.Cell(currentRow, 1).IsEmpty())
                    {
                        currentRow++;
                    }

                    worksheet.Cell(currentRow, 1).Value = bill.billName;
                    worksheet.Cell(currentRow, 2).Value = bill.amount;
                    worksheet.Cell(currentRow, 3).FormulaA1 = bill.isSplit ? $"B{currentRow}/2" : $"B{currentRow}";
                    worksheet.Cell(currentRow, 4).Value = bill.week;
                    worksheet.Cell(currentRow, 5).FormulaA1 = $"D{currentRow}-1";
                    worksheet.Cell(currentRow, 6).FormulaA1 = $"IF(E{currentRow}=0,1,D{currentRow}*7-7)";
                    worksheet.Cell(currentRow, 7).Value = bill.autopayStatus;
                    worksheet.Cell(currentRow, 8).FormulaA1 = $"IF(I{currentRow}<>0,\"Y\",\"N\")";

                    currentRow++;
                }
                workbook.SaveAs(workbookFileName);
                Console.WriteLine($"Template file created at {workbookFileName}");
            }
        }*/
        static void FinalizeCurrentSheet(string workbookFileName)
        {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                if (!workbook.Worksheets.Any())
                {
                    Console.WriteLine($"No worksheets found in the workbook.");
                    return;
                }

                var currentSheet = workbook.Worksheets.First();

                int lastRow = currentSheet.LastRowUsed().RowNumber();

               /* bool hasCompletedSection = true;
                for (int row = 1; row <= lastRow; row++)
                {
                    var cellA = currentSheet.Cell(row, 1);
                    var cellI = currentSheet.Cell(row, 9); // Prob should only check for data in Col. I; May be more efficient to check one but is likely more error proof to check both
                    if (!string.IsNullOrEmpty(cellA.Value.ToString()) && string.IsNullOrEmpty(cellI.Value.ToString()))
                    {
                        hasCompletedSection = false;
                        break;
                    }
                }
                //FORMAT HEADERS SOMEWHERE IN HERE
                if (hasCompletedSection)
                {*/
                    Console.Write("Enter the month to use as the name of the sheet you're saving: ");
                    string newSheetName = Console.ReadLine().Trim();

                    var newSheet = currentSheet.CopyTo(newSheetName);

                    newSheet.Name = newSheetName;

                    for (int row = 2; row <= currentSheet.LastRowUsed().RowNumber(); row++)
                    {
                        currentSheet.Cell(row, 9).Clear();
                    }

                    Console.Write("Do you want to add more bills for the current data entry sheet for your new month? (yes/no): ");
                    string addMoreBills = Console.ReadLine().Trim().ToLower();

                    if (addMoreBills == "yes")
                    {
                        AddBillsToCurrentSheet(workbookFileName);
                    }

                    workbook.Save();
                    Console.WriteLine($"Current sheet finalized and a new sheet named '{newSheetName}' for data entry has been created.");
                /*}
                else
                {
                    Console.WriteLine($"Current sheet does not have a completed section.");
                }*/
            }




        }
        static void AddBillsToCurrentSheet(string workbookFileName)
        {
            /*if (!File.Exists(workbookFileName))
            {
                Console.WriteLine($"No workbook file found for the current year.");
                return;
            }*/
            var existingBills = new Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>>();
            var bills = new List<(string billName, decimal amount, bool isSplit, int week, string autopayStatus)>();
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                var worksheet = workbook.Worksheets.First();
                int currentWeek = 0;

                for (int row = 2; row <= worksheet.LastRowUsed().RowNumber(); row++)
                {
                    string cellValue = worksheet.Cell(row, 1).GetString();
                    if (cellValue.StartsWith("Week "))
                    {
                        currentWeek = int.Parse(cellValue.Replace("Week ", ""));
                        if (!existingBills.ContainsKey(currentWeek))
                        {
                            existingBills[currentWeek] = new List<(string billName, decimal amount, bool isSplit, string autopayStatus)>();
                        }
                    }
                    else if (!string.IsNullOrEmpty(cellValue) && currentWeek != 0)
                    {
                        existingBills[currentWeek].Add((
                            cellValue,
                            worksheet.Cell(row, 2).GetValue<decimal>(),
                            worksheet.Cell(row, 3).FormulaA1.Contains("/2"),
                            worksheet.Cell(row, 7).GetString()
                        ));
                    }
                }

                for (int week = 1; week <= 4; week++)
                {
                    string weekString = $"Week {week}";
                    Console.WriteLine($"Enter bills for {weekString}:");
                    Console.Write($"How many bills do you have for {weekString}? ");
                    int numberOfBills = int.Parse(Console.ReadLine());

                    for (int i = 0; i < numberOfBills; i++)
                    {
                        Console.Write($"Enter the name of bill {i + 1} for {weekString}: ");
                        string billName = Console.ReadLine();
                        Console.Write($"Enter the amount for {billName}: ");
                        decimal amount = decimal.Parse(Console.ReadLine());
                        Console.Write($"Are you splitting {billName} with a roommate? (yes/no): ");
                        bool isSplit = Console.ReadLine().Trim().ToLower() == "yes";
                        Console.Write($"Enter autopay status for {billName} (yes/no): ");
                        string autopayStatus = Console.ReadLine().Trim().ToLower();

                        if (!existingBills.ContainsKey(week))
                        {
                            existingBills[week] = new List<(string billName, decimal amount, bool isSplit, string autopayStatus)>();
                        }
                        existingBills[week].Add((billName, amount, isSplit, autopayStatus));
                    }
                }

                worksheet.Clear();

                //Format headers
                worksheet.Cell("A1").Value = "Bill Name";
                worksheet.Cell("B1").Value = "Minimum Amount Owed";
                worksheet.Cell("C1").Value = "Minimum Amount Due";
                worksheet.Cell("D1").Value = "Due Date Week";
                worksheet.Cell("E1").Value = "Transition Formula";
                worksheet.Cell("F1").Value = "Latest Due Date";
                worksheet.Cell("G1").Value = "Autopay Status";
                worksheet.Cell("H1").Value = "Paid Boolean";
                worksheet.Cell("I1").Value = "Amount Paid";
                //probably bad logic
                int currentRow = 2;
                for (int week = 1; week <= 4; week++)
                {
                    worksheet.Cell(currentRow, 1).Value = $"Week {week}";
                    currentRow++;

                    if (existingBills.ContainsKey(week))
                    {
                        foreach (var bill in existingBills[week])
                        {
                            worksheet.Cell(currentRow, 1).Value = bill.billName;
                            worksheet.Cell(currentRow, 2).Value = bill.amount;
                            worksheet.Cell(currentRow, 3).FormulaA1 = bill.isSplit ? $"B{currentRow}/2" : $"B{currentRow}";
                            worksheet.Cell(currentRow, 4).Value = week;
                            worksheet.Cell(currentRow, 5).FormulaA1 = $"D{currentRow}-1";
                            worksheet.Cell(currentRow, 6).FormulaA1 = $"IF(E{currentRow}=0,1,D{currentRow}*7-7)";
                            worksheet.Cell(currentRow, 7).Value = bill.autopayStatus;
                            worksheet.Cell(currentRow, 8).FormulaA1 = $"IF(I{currentRow}<>0,\"Y\",\"N\")";
                            currentRow++;
                        }
                    }
                }

                // Save the workbook to persist the changes
                workbook.Save();

                // Format the worksheet
                var headerRange = worksheet.Range("A1:E1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                foreach (var column in worksheet.Columns())
                {
                    column.Width = 26;
                }

                workbook.Save();
                Console.WriteLine("New bills added to the current worksheet.");
                //IF YOU GOT TO THE END OF THIS, MY HANDS HURT
            }
        }
    }
}
