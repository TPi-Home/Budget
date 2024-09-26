//6/13/24 Tyler Pittman
using ClosedXML.Excel;

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
                //determine cell bg colors
                var headerColor = XLColor.Black;
                var bodyColor = XLColor.FromHtml("#d19ffc");
                var weekRowColor = XLColor.FromHtml("#4f1c75");
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
                var lastRow = currentSheet.LastRowUsed()?.RowNumber() ?? 0;
                //apply body theme
                var bodyRows = range.Rows(row => row.RowNumber() > headerRow.RowNumber());//check for redundant variables
                for (int row = 2; row <= lastRow; row++)
                {
                    string cellValue = currentSheet.Cell(row, 1).GetString();
                    if (cellValue.StartsWith("Week ") || cellValue.StartsWith("Total"))
                    {
                        //if it's a divider row:
                        for (int col = 1; col <= currentSheet.LastColumnUsed().ColumnNumber(); col++)
                        {
                            currentSheet.Cell(row, col).Style.Fill.BackgroundColor = weekRowColor;
                            currentSheet.Cell(row, col).Style.Font.FontColor = XLColor.White;
                            //border:
                            currentSheet.Cell(row, col).Style.Border.SetTopBorder(XLBorderStyleValues.Medium);
                            currentSheet.Cell(row, col).Style.Border.SetRightBorder(XLBorderStyleValues.Medium);
                            currentSheet.Cell(row, col).Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);
                            currentSheet.Cell(row, col).Style.Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            currentSheet.Cell(row, col).Style.Font.Bold = true;
                        }
                    }
                    //else it's a body row:
                    else
                    {
                        for (int col = 1; col <= currentSheet.LastColumnUsed().ColumnNumber(); col++)
                        {
                            currentSheet.Cell(row, col).Style.Fill.BackgroundColor = bodyColor;
                            currentSheet.Cell(row, col).Style.Font.FontColor = XLColor.Black;
                            //border:
                            currentSheet.Cell(row, col).Style.Border.SetTopBorder(XLBorderStyleValues.Medium);
                            currentSheet.Cell(row, col).Style.Border.SetRightBorder(XLBorderStyleValues.Medium);
                            currentSheet.Cell(row, col).Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);
                            currentSheet.Cell(row, col).Style.Border.SetLeftBorder(XLBorderStyleValues.Medium);
                            if (col == 1)
                                currentSheet.Cell(row, col).Style.Font.Bold = true;
                        }
                    }
                }
                workbook.Save();
                Console.WriteLine("Workbook saved successfully with applied theme.");
            }
        }

        public static void applyHeaders(string workbookFileName)
        {
            using (var workbook = new XLWorkbook(workbookFileName))
            {
                var currentSheet = workbook.Worksheets.FirstOrDefault();
                //formatting headers
                //these should be variables in case of user defined columns 
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
                    column.Width = 22;//as conditions imply
                }
                workbook.Save();
            }
        }
    }
}
