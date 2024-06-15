//6/13/24 Tyler Pittman
using ClosedXML.Excel;

namespace Budget
{
    internal class WSAppend //Definitely too much going on in this class.
    {
        public static void AddBillsToCurrentSheet(string workbookFileName, Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> existingBills)
        {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                var currentSheet = workbook.Worksheets.FirstOrDefault();
                if (currentSheet == null)
                {
                    throw new InvalidOperationException("The workbook must contain at least one worksheet.");
                }
                var lastRow = currentSheet.LastRowUsed()?.RowNumber() ?? 0;
                int currentWeek = 0;
                for (int row = 1; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue.StartsWith("Week "))
                    {
                        currentWeek = int.Parse(cellValue.Replace("Week ", ""));
                        break;
                    }
                }
                //bool totalExists = false;
                for (int row = 1; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue == "Total:")
                    {
                        //totalExists = true;
                        lastRow--;
                        break;
                    }
                }
                for (int row = 2; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
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
                        //check for duplicates
                        if (!existingBills[currentWeek].Exists(bill => bill.billName == cellValue))
                        {
                            existingBills[currentWeek].Add((
                                cellValue,
                                currentSheet.Cell(row, 2).GetValue<decimal>(),
                                currentSheet.Cell(row, 3).FormulaA1.Contains("/2"),
                                currentSheet.Cell(row, 7).GetString()
                            ));
                        }
                    }
                }
                //dataGather(workbookFileName, existingBills);
                DataGather(workbookFileName,existingBills);
                //print the collected bills for verification
                foreach (var weekBills in existingBills)
                {
                    Console.WriteLine($"Bills for Week {weekBills.Key}:");
                    foreach (var bill in weekBills.Value)
                    {
                        Console.WriteLine($"- {bill.billName}: {bill.amount}, Split: {bill.isSplit}, Autopay: {bill.autopayStatus}");
                    }
                }
                currentSheet.Clear();
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
                            currentSheet.Cell(currentRow, 3).FormulaA1 = bill.isSplit
                                ? $"=IF(H{currentRow}=\"Y\",IF(B{currentRow}/2-I{currentRow}<0,0,B{currentRow}/2-I{currentRow}),B{currentRow}/2)"
                                : $"=IF(H{currentRow}=\"Y\",IF(B{currentRow}-I{currentRow}<0,0,B{currentRow}-I{currentRow}),B{currentRow})";
                            currentSheet.Cell(currentRow, 4).Value = week;
                            currentSheet.Cell(currentRow, 5).FormulaA1 = $"D{currentRow}-1";
                            currentSheet.Cell(currentRow, 6).FormulaA1 = $"IF(E{currentRow}=0,1,D{currentRow}*7-7)";
                            currentSheet.Cell(currentRow, 7).Value = bill.autopayStatus;
                            currentSheet.Cell(currentRow, 8).FormulaA1 = $"IF(I{currentRow}<>0,\"Y\",\"N\")";
                            currentRow++;
                        }
                    }
                }
                workbook.Save();
                Console.WriteLine("New bills added to the current worksheet.");
            }
        }

        private static void DataGather(string workbookFileName, Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> existingBills)//make class and break up into helper functions
        {
            for (int week = 1; week <= 4; week++)
            {
                string weekString = $"Week {week}";
                int numberOfBills;

                while (true)
                {
                    Console.Write($"How many new bills do you have for {weekString}? ");//enter a break after 
                    string? input = Console.ReadLine();

                    if (int.TryParse(input, out numberOfBills) && numberOfBills >= 0)
                    {
                        break;
                    }
                    else
                    {
                        Console.WriteLine("Invalid input. Please enter a non-negative integer.");
                    }
                }

                for (int i = 0; i < numberOfBills; i++)
                {
                    string? billName;
                    while (true)
                    {
                        Console.Write($"Enter the name of bill {i + 1} for {weekString}: ");
                        billName = Console.ReadLine();

                        if (!string.IsNullOrWhiteSpace(billName))
                        {
                            break;
                        }
                        else
                        {
                            Console.WriteLine("Invalid input. Bill name cannot be empty.");
                        }
                    }

                    decimal amount;
                    while (true)
                    {
                        Console.Write($"Enter the amount for {billName}: ");
                        string? input = Console.ReadLine();

                        if (decimal.TryParse(input, out amount) && amount >= 0)
                        {
                            break;
                        }
                        else
                        {
                            Console.WriteLine("Invalid input. Please enter a non-negative decimal number.");
                        }
                    }

                    bool isSplit;
                    while (true)
                    {
                        Console.Write($"Are you splitting {billName} with a roommate? (yes/no): ");
                        string? input = Console.ReadLine().Trim().ToLower();

                        if (input == "yes")
                        {
                            isSplit = true;
                            break;
                        }
                        else if (input == "no")
                        {
                            isSplit = false;
                            break;
                        }
                        else
                        {
                            Console.WriteLine("Invalid input. Please enter 'yes' or 'no'.");
                        }
                    }

                    string? autoPayStatus;
                    while (true)
                    {
                        Console.Write($"Enter autopay status for {billName} (yes/no): ");
                        autoPayStatus = Console.ReadLine().Trim().ToLower();

                        if (autoPayStatus == "yes" || autoPayStatus == "no")
                        {
                            break;
                        }
                        else
                        {
                            Console.WriteLine("Invalid input. Please enter 'yes' or 'no'.");
                        }
                    }

                    if (!existingBills.ContainsKey(week))
                    {
                        existingBills[week] = new List<(string billName, decimal amount, bool isSplit, string autopayStatus)>();
                    }

                    if (!existingBills[week].Exists(bill => bill.billName == billName))
                    {
                        existingBills[week].Add((billName, amount, isSplit, autoPayStatus));
                    }
                    else
                    {
                        Console.WriteLine($"Bill with the name {billName} already exists for {weekString}. Skipping duplicate.");
                    }
                }
            }
        }
        public static void TotalsAndFormula(string workbookFileName)
        {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                var currentSheet = workbook.Worksheets.First();
                var lastRow = currentSheet.LastRowUsed()?.RowNumber() ?? 0;
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
                        break;
                    }
                }
                if (!totalExists)
                {
                    lastRow++;
                    currentSheet.Cell(lastRow, 1).Value = "Total:";
                }
                workbook.Save();
                Dictionary<int, string> weekTotalFormulas = new Dictionary<int, string>();
                List<string> weekTotalCells = new List<string>();
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
                        if (weekTotalStartRow <= weekTotalEndRow - 1)
                        {
                            string weekTotalFormula = $"SUM(C{weekTotalStartRow}:C{weekTotalEndRow - 1})";
                            weekTotalFormulas[weekNumber] = weekTotalFormula;
                            currentSheet.Cell(row, 3).FormulaA1 = weekTotalFormula;
                            weekTotalCells.Add($"C{row}");
                            //paid total here
                        }
                    }
                }
                string monthTotalFormula = $"SUM({string.Join(",", weekTotalCells)})";
                currentSheet.Cell(lastRow, 3).FormulaA1 = monthTotalFormula;


                workbook.Save();
            }
        }
        public static void FinalizeCurrentSheet(string workbookFileName, Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> existingBills)
        {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                if (!workbook.Worksheets.Any())
                {
                    Console.WriteLine($"No worksheets found in the workbook.");
                    return;
                }
                var currentSheet = workbook.Worksheets.First();//maybe prompt user about name. could .ToUpper the sheet names and compare
                int lastRow = currentSheet.LastRowUsed().RowNumber();
                Console.Write("Enter the month to use as the name of the sheet you're saving: ");
                string newSheetName = Console.ReadLine().Trim();
                var newSheet = currentSheet.CopyTo(newSheetName);
                newSheet.Name = newSheetName;
                for (int row = 2; row <= currentSheet.LastRowUsed().RowNumber(); row++)
                {
                    currentSheet.Cell(row, 9).Clear();
                }
                //workbook.Worksheet(newSheet).SetTabActive();
                workbook.Save();
                Console.WriteLine($"Current sheet finalized and a new sheet named '{newSheetName}' for data entry has been created.");
            }
        }
    }
}



