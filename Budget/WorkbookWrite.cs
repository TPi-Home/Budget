using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;


namespace Budget
{
    internal class WorkbookWrite
    {
        public static WorkbookResult CreateBlank(string workbookFileName, Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> existingBills)
        {
            if (File.Exists(workbookFileName))
            {
                return new WorkbookResult
                {
                    Success = false,
                    Message = "You already have a Workbook for this year!"
                };
            }
            Console.Write("Enter the month to use as the name of the sheet you're saving: ");
            string newSheetName = Console.ReadLine().Trim();
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    workbook.AddWorksheet(newSheetName);  //ensure at least one worksheet is added
                    workbook.SaveAs(workbookFileName);
                }
                return new WorkbookResult
                {
                    Success = true,
                    Message = $"Workbook created: {workbookFileName}"
                };
            }
            catch (Exception ex)
            {
                return new WorkbookResult
                {
                    Success = false,
                    Message = $"An error occurred: {ex.Message}"
                };
            }
        }

    }
}
    

