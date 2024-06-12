using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Budget
{
    internal class WS //May be too much going on in this class.
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

                //adding new bills for each week
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

                        //check for duplicates again. may be excessive
                        if (!existingBills[week].Exists(bill => bill.billName == billName))
                        {
                            existingBills[week].Add((billName, amount, isSplit, autopayStatus));
                        }
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
                workbook.Save();//probably also redundant?
                //format the worksheet
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
        public static void totalsAndFormula(string workbookFileName)
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
    }
}


