using ClosedXML.Excel;
using System;

namespace Budget
{
    internal class TableTheme
    {
        public void applyTheme(string workbookFileName)
        {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                var currentSheet = workbook.Worksheets.FirstOrDefault();
                if (currentSheet == null)
                {
                    Console.WriteLine("No worksheet found in the workbook.");
                    return;
                }

                //determine the range based on used cells
                var firstCell = currentSheet.FirstCellUsed();
                var lastCell = currentSheet.LastCellUsed();
                if (firstCell == null || lastCell == null)
                {
                    Console.WriteLine("The worksheet is empty.");
                    return;
                }

                var headerColor = XLColor.CoolBlack;
                var bodyColor = XLColor.AliceBlue;

                var range = currentSheet.Range(firstCell.Address, lastCell.Address);
                if (range == null)
                {
                    Console.WriteLine("Failed to determine the range.");
                    return;
                }

                //apply header theme
                var headerRow = range.FirstRow();
                if (headerRow != null)
                {
                    headerRow.Style.Fill.BackgroundColor = headerColor;
                    headerRow.Style.Font.Bold = true;
                    headerRow.Style.Font.FontColor = XLColor.White;
                    headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                else
                {
                    Console.WriteLine("Failed to find the header row.");
                }

                //apply body theme
                var bodyRows = range.Rows(row => row.RowNumber() > headerRow.RowNumber());
                foreach (var row in bodyRows)
                {
                    row.Style.Fill.BackgroundColor = bodyColor;
                }

                workbook.Save();
                Console.WriteLine("Workbook saved successfully with applied theme.");
            }
        }
        
        public static void applyHeaders(string workbookFileName) {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                var currentSheet = workbook.Worksheets.FirstOrDefault();
                //formatting headers
                currentSheet.Cell("A1").Value = "Bill Name";
                currentSheet.Cell("B1").Value = "Minimum Amount Owed";
                currentSheet.Cell("C1").Value = "Minimum Amount Due";
                currentSheet.Cell("D1").Value = "Due Date Week";
                currentSheet.Cell("E1").Value = "Transition Formula";
                currentSheet.Cell("F1").Value = "Latest Due Date";
                currentSheet.Cell("G1").Value = "Autopay Status";
                currentSheet.Cell("H1").Value = "Paid Boolean";
                currentSheet.Cell("I1").Value = "Amount Paid";

                var headerRange = currentSheet.Range("A1:I1");
                headerRange.Style.Font.Bold = true;
                //headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                foreach (var column in currentSheet.Columns())
                {
                    column.Width = 26;//as conditions imply
                }
                workbook.Save();
            }
        }
    }
}
