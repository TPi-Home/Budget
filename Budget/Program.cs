using ClosedXML.Excel;
//deleted using openxml

namespace BudgetSpreadsheet
{
    class Program
    {
        static void Main(string[] args)
        {
            string currentYear = DateTime.Now.ToString("yyyy");
            string workbookFileName = $"{currentYear}.xlsx";

            var existingBills = new Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>>();

            // Initial Menu
            Console.WriteLine("What do you want to do?");
            Console.WriteLine("1. Create an appropriately named blank spreadsheet in the root directory.\n(This software follows specific naming conventions, such as the current year in the case of the workbook)");
            Console.WriteLine("2. Add and merge bills into an existing worksheet.");
            Console.WriteLine("3. Add formulas to final cells to determine to weekly and monthly expense totals.");//should probably clarify what this means; finalize month
            Console.WriteLine("4. Finalize the current sheet and start a new sheet for data entry.");
            Console.WriteLine("5. Open tutorial.");
            Console.WriteLine("6. Open README.");
            Console.WriteLine("7. Exit.");
            Console.Write("Enter your choice (1, 2, 3, 4, 5, 6 or 7): ");//MUST ADD INPUT VALIDTION
            string choice = Console.ReadLine();
            while (choice != "7")
            {
                switch (choice)
                {
                    case "1":
                        //create blank file
                        createBlank(workbookFileName, existingBills);
                        break;
                    case "2":
                        //create a custom template sheet
                        AddBillsToCurrentSheet(workbookFileName, existingBills);
                        break;
                    case "3":
                        //add total formulas and final touchups; color code
                        totalsAndFormula(workbookFileName);
                        break;
                    case "4":
                        //finalize
                        FinalizeCurrentSheet(workbookFileName, existingBills);
                        break;
                    case "5":
                        tutorial();
                        break;
                    case "6":
                        readMe();
                        break;
                    case "7":
                        Console.WriteLine("Closing.");
                        break;
                    default:
                        Console.WriteLine("Invalid choice.");
                        break;
                }
                Console.WriteLine("What do you want to do?");
                Console.WriteLine("1. Create an appropriately named blank spreadsheet in the root directory.\n(This software follows specific naming conventions, such as the current year in the case of the workbook)");
                Console.WriteLine("2. Add and merge bills into an existing worksheet.");//will be adding a remove option later
                Console.WriteLine("3. Add formulas to final cells to determine to weekly and monthly expense totals.");
                Console.WriteLine("4. Finalize the current sheet and start a new sheet for data entry.");//should probably clarify what this means; finalize month
                Console.WriteLine("5. Open tutorial.");
                Console.WriteLine("6. Open README.");
                Console.WriteLine("7. Exit.");
                Console.Write("Enter your choice (1, 2, 3, 4, 5, 6 or 7): ");

                //break up rent and mortgage, taxes (plus adjusting applicable bills to tax rate, prob seperate column), seperate insurance types, subscription audit, income and capital gains, 1099 income, savings info, debt payments, cells for tax season reminders
                //add delete expensese, fix logic, handle exceptions
                //check for open file
                
                choice = Console.ReadLine();
            }

        }
        static void createBlank(string workbookFileName, Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> existingBills)
        {
            using (var workbook = new XLWorkbook())
            {
                int currentWeek = 0;
                string[] args = null;
                //check for file in root directory
                if (File.Exists(workbookFileName))
                {
                    Console.WriteLine($"You already have a Workbook for this year!");
                    return;
                }
                var worksheet = workbook.Worksheets.Add("Entry");
                var currentSheet = workbook.Worksheets.First();

                for (int week = 1; week <= 4; week++)
                {
                    string weekString = $"Week {week}";
                    Console.Write($"How many bills do you have for {weekString}? ");
                    int numberOfBills = int.Parse(Console.ReadLine());

                    for (int i = 0; i < numberOfBills; i++)
                    {
                        Console.Write($"Enter the name of bill {i + 1} for {weekString}: ");
                        string billName = Console.ReadLine();
                        Console.Write($"Enter the amount for {billName}: ");
                        decimal amount = decimal.Parse(Console.ReadLine());
                        Console.Write($"Are you splitting {billName} with a roommate? (yes/no): ");
                        bool isSplit = Console.ReadLine().Trim().ToLower().StartsWith("y");
                        Console.Write($"Enter autopay status for {billName} (yes/no): ");
                        string autopayStatus = Console.ReadLine().Trim().ToLower();

                        if (!existingBills.ContainsKey(week))
                        {
                            existingBills[week] = new List<(string billName, decimal amount, bool isSplit, string autopayStatus)>();
                        }
                        existingBills[week].Add((billName, amount, isSplit, autopayStatus));
                    }
                }

                currentSheet.Clear();

                //formatting
                currentSheet.Cell("A1").Value = "Bill Name";
                currentSheet.Cell("B1").Value = "Minimum Amount Owed";
                currentSheet.Cell("C1").Value = "Minimum Amount Due";
                currentSheet.Cell("D1").Value = "Due Date Week";
                currentSheet.Cell("E1").Value = "Transition Formula";
                currentSheet.Cell("F1").Value = "Latest Due Date";
                currentSheet.Cell("G1").Value = "Autopay Status";
                currentSheet.Cell("H1").Value = "Paid Boolean";
                currentSheet.Cell("I1").Value = "Amount Paid";
                //probably bad logic
                int currentRow = 2;
                for (int week = 1; week <= 4; week++)
                {
                    currentSheet.Cell(currentRow, 1).Value = $"Week {week}";
                    currentRow++;

                    if (existingBills.ContainsKey(week))
                    {
                        foreach (var bill in existingBills[week])
                        {
                            currentSheet.Cell(currentRow, 1).Value = bill.billName;
                            currentSheet.Cell(currentRow, 2).Value = bill.amount;
                            if (bill.isSplit)
                                currentSheet.Cell(currentRow, 3).FormulaA1 = $"=IF(H{currentRow}=\"Y\",IF(B{currentRow}/2-I{currentRow}<0,0,B{currentRow}/2-I{currentRow}),B{currentRow}/2)";
                            else
                                currentSheet.Cell(currentRow, 3).FormulaA1 = $"=IF(H{currentRow}=\"Y\",IF(B{currentRow}-I{currentRow}<0,0,B{currentRow}-I{currentRow}),B{currentRow})";
                            currentSheet.Cell(currentRow, 4).Value = week;
                            currentSheet.Cell(currentRow, 5).FormulaA1 = $"D{currentRow}-1";
                            currentSheet.Cell(currentRow, 6).FormulaA1 = $"IF(E{currentRow}=0,1,D{currentRow}*7-7)";
                            currentSheet.Cell(currentRow, 7).Value = bill.autopayStatus;
                            currentSheet.Cell(currentRow, 8).FormulaA1 = $"IF(I{currentRow}<>0,\"Y\",\"N\")";
                            currentRow++;
                        }
                    }
                }

                // Save the workbook to persist the changes
                workbook.SaveAs(workbookFileName);

                // Format the worksheet
                var headerRange = currentSheet.Range("A1:I1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                foreach (var column in currentSheet.Columns())
                {
                    column.Width = 26;
                }

                workbook.Save();
                Console.WriteLine("New bills added to the current worksheet.");
            }
        }
        static void totalsAndFormula(string workbookFileName)
        {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                var currentSheet = workbook.Worksheets.First();
                var lastRow = currentSheet.LastRowUsed().RowNumber();
                //month total and final row logic
                int weekEndRow = 1;
                for (int row = 1; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue.StartsWith("Week"))
                    {
                        weekEndRow = row - 1;
                        break;
                    }
                }
                bool totalExists = false;
                for (int row = 1; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue == "Total:")
                    {
                        totalExists = true;
                        //lastRow--;
                        break;
                    }
                }

                if (!totalExists)
                {
                    lastRow++;
                    currentSheet.Cell(lastRow, 1).Value = "Total:";

                }

                workbook.Save();
                // Dictionary to store week totals
                Dictionary<int, string> weekTotalFormulas = new Dictionary<int, string>();
                List<string> weekTotalCells = new List<string>();
                // Iterate through the rows to calculate week totals
                for (int row = 2; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue.StartsWith("Week "))
                    {
                        int weekNumber = int.Parse(cellValue.Replace("Week ", ""));
                        int weekTotalStartRow = row + 1;
                        int weekTotalEndRow = weekTotalStartRow;

                        while (weekTotalEndRow <= lastRow && !currentSheet.Cell(weekTotalEndRow, 1).GetString().StartsWith("Week ") && !currentSheet.Cell(weekTotalEndRow, 1).GetString().Equals("Total:"))
                        {
                            weekTotalEndRow++;
                        }

                        // Check if there are bills for the given week
                        if (weekTotalStartRow <= weekTotalEndRow - 1)
                        {
                            string weekTotalFormula = $"SUM(C{weekTotalStartRow}:C{weekTotalEndRow - 1})";
                            weekTotalFormulas[weekNumber] = weekTotalFormula;
                            currentSheet.Cell(row, 3).FormulaA1 = weekTotalFormula;
                            weekTotalCells.Add($"C{row}");
                        }
                    }
                }

                string monthTotalFormula = $"SUM({string.Join(",", weekTotalCells)})";
                currentSheet.Cell(lastRow, 3).FormulaA1 = monthTotalFormula;
                // Save the workbook
                workbook.Save();
            }

          
        }
        static void tutorial()
        {
            Console.Write("To use this software, enter a number from the main menu and press enter. Use option 1 to create a spreadsheet in the same directory this software is running. Follow the prompts to have accurate expense information in your new budget spreadsheet. \nTo add new or forgotten expenses, use option number 2 from the main menu.\n" +
                "Option number 3 on the main menu cleans up and finalizes much of the template with additional formatting and adds formulas for weekly and monthly totals.\n" +
                "Once you are done with a month, use option 4 if you would like to save your sheet with paid amounts filled out to a new sheet and have cleared data entry sheet. To use the spreadsheet, simply fill out your amount paid using the cells in column I. The rest of the cells should populate based off the data entered.");
        }
        static void readMe()
        {
            Console.WriteLine("*This software gathers information from the user and builds a spreadsheet based off that information.\n*This is the style of spreadsheet I have used to track my expenses/budget over the last couple years.\n*While I have been content with this model for my personal use, I do believe there will be more to come as I clean it up some.\n" +
                "Please report bugs directly to me if you know me\n \n \n \nCopyright 2024 Tyler Pittman\r\n\r\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\r\n\r\nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\r\n\r\nTHE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. ");
        }

        static void FinalizeCurrentSheet(string workbookFileName, Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> existingBills)
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

                    Console.Write("Enter the month to use as the name of the sheet you're saving: ");
                    string newSheetName = Console.ReadLine().Trim();

                    var newSheet = currentSheet.CopyTo(newSheetName);

                    newSheet.Name = newSheetName;

                    for (int row = 2; row <= currentSheet.LastRowUsed().RowNumber(); row++)
                    {
                        currentSheet.Cell(row, 9).Clear();
                    }
                    workbook.Save();
                    Console.WriteLine($"Current sheet finalized and a new sheet named '{newSheetName}' for data entry has been created.");
            }
        }
        static void AddBillsToCurrentSheet(string workbookFileName, Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> existingBills)
        {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                var currentSheet = workbook.Worksheets.First();
                var lastRow = currentSheet.LastRowUsed().RowNumber();
                int currentWeek = 0;
                int weekEndRow = 1;
                for (int row = 1; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue.StartsWith("Week"))
                    {
                        weekEndRow = row - 1;
                        break;
                    }
                }
                bool totalExists = false;
                for (int row = 1; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue == "Total:")
                    {
                        totalExists = true;
                        lastRow--;
                        break;
                    }
                }
                //reading and adding to dict
                for (int row = 2; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue.StartsWith("Week "))
                    {//user needs to keep placeholder cells in A column indicating week if they want to use this software as this is critical to being able to sort in this context
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
                            currentSheet.Cell(row, 2).GetValue<decimal>(),
                            currentSheet.Cell(row, 3).FormulaA1.Contains("/2"),
                            currentSheet.Cell(row, 7).GetString()
                        ));
                    }
                }

                for (int week = 1; week <= 4; week++)
                {
                    string weekString = $"Week {week}";
                    Console.Write($"How many new bills do you have for {weekString}? ");
                    int numberOfBills = int.Parse(Console.ReadLine());

                    for (int i = 0; i < numberOfBills; i++)
                    {
                        Console.Write($"Enter the name of bill {i + 1} for {weekString}: ");
                        string billName = Console.ReadLine();
                        Console.Write($"Enter the amount for {billName}: ");
                        decimal amount = decimal.Parse(Console.ReadLine());
                        Console.Write($"Are you splitting {billName} with a roommate? (yes/no): ");
                        bool isSplit = Console.ReadLine().Trim().ToLower().StartsWith("y");
                        Console.Write($"Enter autopay status for {billName} (yes/no): ");
                        string autopayStatus = Console.ReadLine().Trim().ToLower();

                        if (!existingBills.ContainsKey(week))
                        {
                            existingBills[week] = new List<(string billName, decimal amount, bool isSplit, string autopayStatus)>();
                        }
                        existingBills[week].Add((billName, amount, isSplit, autopayStatus));
                    }
                }

                currentSheet.Clear();

                //formatting
                currentSheet.Cell("A1").Value = "Bill Name";
                currentSheet.Cell("B1").Value = "Minimum Amount Owed";
                currentSheet.Cell("C1").Value = "Minimum Amount Due";
                currentSheet.Cell("D1").Value = "Due Date Week";
                currentSheet.Cell("E1").Value = "Transition Formula";
                currentSheet.Cell("F1").Value = "Latest Due Date";
                currentSheet.Cell("G1").Value = "Autopay Status";
                currentSheet.Cell("H1").Value = "Paid Boolean";
                currentSheet.Cell("I1").Value = "Amount Paid";
                //probably bad logic
                int currentRow = 2;
                for (int week = 1; week <= 4; week++)
                {
                    currentSheet.Cell(currentRow, 1).Value = $"Week {week}";
                    currentRow++;

                    if (existingBills.ContainsKey(week))
                    {
                        foreach (var bill in existingBills[week])
                        {
                            currentSheet.Cell(currentRow, 1).Value = bill.billName;
                            currentSheet.Cell(currentRow, 2).Value = bill.amount;
                            if (bill.isSplit==true)
                                currentSheet.Cell(currentRow, 3).FormulaA1 = $"=IF(H{currentRow}=\"Y\",IF(B{currentRow}/2-I{currentRow}<0,0,B{currentRow}/2-I{currentRow}),B{currentRow}/2)";
                            else
                                currentSheet.Cell(currentRow, 3).FormulaA1 = $"=IF(H{currentRow}=\"Y\",IF(B{currentRow}-I{currentRow}<0,0,B{currentRow}-I{currentRow}),B{currentRow})";
                            currentSheet.Cell(currentRow, 4).Value = week;
                            currentSheet.Cell(currentRow, 5).FormulaA1 = $"D{currentRow}-1";
                            currentSheet.Cell(currentRow, 6).FormulaA1 = $"IF(E{currentRow}=0,1,D{currentRow}*7-7)";
                            currentSheet.Cell(currentRow, 7).Value = bill.autopayStatus;
                            currentSheet.Cell(currentRow, 8).FormulaA1 = $"IF(I{currentRow}<>0,\"Y\",\"N\")";
                            currentRow++;
                        }

                    }
                }

                // Save the workbook to persist the changes
                workbook.Save();

                // Format the worksheet
                var headerRange = currentSheet.Range("A1:I1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                foreach (var column in currentSheet.Columns())
                {
                    column.Width = 26;
                }

                workbook.Save();
                Console.WriteLine("New bills added to the current worksheet.");
                existingBills = null;
                //IF YOU GOT TO THE END OF THIS, MY HANDS HURT
                //debug to see if bool totalexists or whatever needs reset
            }
        }
    }
}
