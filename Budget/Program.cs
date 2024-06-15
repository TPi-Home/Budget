//6/13/24 Tyler Pittman
using Budget;

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
            Console.WriteLine("2. Add or merge bills into default worksheet.");
            Console.WriteLine("3. Add formulas to final cells and apply table theme to default sheet.");
            Console.WriteLine("4. Finalize the current sheet and start a new sheet for data entry.");
            Console.WriteLine("5. Open tutorial.");
            Console.WriteLine("6. Open README.");
            Console.WriteLine("7. Exit.");
            Console.Write("Enter your choice (1, 2, 3, 4, 5, 6 or 7): ");
            string? choice = Console.ReadLine();
            /*Random choice = new Random(); 
            var unitTest = choice.Next(1, 4);
            string nchoice = unitTest.ToString();*/
            while (choice != "7")
            {
                switch (choice)
                {
                    case "1":
                        var result = WorkbookWrite.CreateBlank(workbookFileName, existingBills);

                        if (result.Success)
                        {
                            Console.WriteLine(result.Message);
                        }
                        else
                        {
                            Console.WriteLine($"Failed to create workbook: {result.Message}");
                        }
                        break;
                    case "2":
                        WSAppend.AddBillsToCurrentSheet(workbookFileName, existingBills);//broke here 6/13/2024. didn't crash but it broke the conditionals here
                        break;
                    case "3":
                        //add total formulas and final touchups; color code
                        WSAppend.TotalsAndFormula(workbookFileName);
                        //var worksheet = workbook.Worksheet(currentSheet);
                        var tableTheme = new TableTheme();

                        // Apply header theme
                        TableTheme.applyHeaders(workbookFileName);

                        // Apply full theme
                        tableTheme.applyTheme(workbookFileName);
                        break;
                    case "4":
                        //finalize
                        WSAppend.FinalizeCurrentSheet(workbookFileName, existingBills);
                        break;
                    case "5":
                        Tutorial();
                        break;//possibly add table themes here
                    case "6":
                        ReadMe();
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
                Console.WriteLine("2. Add or merge bills into default worksheet.");//will be adding a remove option later
                Console.WriteLine("3. Add formulas to final cells and apply table theme to default sheet.");//Are you sure??
                Console.WriteLine("4. Finalize the current sheet and start a new sheet for data entry.");//should probably clarify what this means; finalize month
                Console.WriteLine("5. Open tutorial.");
                Console.WriteLine("6. Open README.");
                Console.WriteLine("7. Exit.");
                Console.Write("Enter your choice (1, 2, 3, 4, 5, 6 or 7): ");// Random rnd = new Random(); int unitTest  = rnd.Next(1, 4);
                //break up rent and mortgage, taxes (plus adjusting applicable bills to tax rate, prob seperate column), seperate insurance types, subscription audit, income and capital gains, 1099 income, savings info, debt payments, cells for tax season reminders
                //add delete expensese
                //check for open file
                //NEED TO ADD AN OPTION TO BREAK EARLY OR MODIFY NUMBER OF BILLS
                //NEED TO BE ABLE TO DYNAMICALLY ADD COLUMNS, ENTER NEW FORMULAS WITH EASE (ML?), AND ADJUST BILL SPLITTING 
                //add sql for storing in program
                //need to add actual budget portion in addition to expense trackings
                //ensure dependencies use MIT license
                choice = Console.ReadLine();
                //unitTest = choice.Next(1, 4);
                //nchoice = unitTest.ToString();
            }
        }
        static void Tutorial()//will probably add some modularity here when I work on the gui
        {
            Console.Write("To use this software, enter a number from the main menu and press enter. Use option 1 to create a spreadsheet in the same directory this software is running. Follow the prompts to have accurate expense information in your new budget spreadsheet. \nTo add new or forgotten expenses, use option number 2 from the main menu.\n" +
                "Option number 3 on the main menu cleans up and finalizes much of the template with additional formatting and adds formulas for weekly and monthly totals.\n" +
                "Once you are done with a month, use option 4 if you would like to save your sheet with paid amounts filled out to a new sheet and have cleared data entry sheet. To use the spreadsheet, simply fill out your amount paid using the cells in column I. The rest of the cells should populate based off the data entered.");
        }
        static void ReadMe()
        {
            Console.WriteLine("*This software gathers information from the user and builds a spreadsheet based off that information.\n*This is the style of spreadsheet I have used to track my expenses/budget over the last couple years.\n*While I have been content with this model for my personal use, I do believe there will be more to come as I clean it up some.\n" +
                "Please report bugs directly to me if you know me\n \n \n \nCopyright 2024 Tyler Pittman\r\n\r\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\r\n\r\nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\r\n\r\nTHE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. ");
        }
    }
}
