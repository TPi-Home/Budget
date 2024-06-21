using System;
using System.Collections.Generic;

namespace Budget
{
    internal class DataGathering
    {
        public static void DataGather(string workbookFileName, Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> existingBills)
        {
            for (int week = 1; week <= 4; week++)
            {
                string weekString = $"Week {week}";
                int numberOfBills = GetNumberOfBills(weekString);

                for (int i = 0; i < numberOfBills; i++)
                {
                    string billName = GetBillName(weekString, i);
                    decimal amount = GetBillAmount(billName);
                    bool isSplit = GetBillSplitStatus(billName);
                    string autopayStatus = GetAutopayStatus(billName);

                    if (!existingBills.ContainsKey(week))
                    {
                        existingBills[week] = new List<(string billName, decimal amount, bool isSplit, string autopayStatus)>();
                    }

                    if (!existingBills[week].Exists(bill => bill.billName == billName))
                    {
                        existingBills[week].Add((billName, amount, isSplit, autopayStatus));
                    }
                    else
                    {
                        Console.WriteLine($"Bill with the name {billName} already exists for {weekString}. Skipping duplicate.");
                    }
                }
            }
        }

        static int GetNumberOfBills(string weekString)
        {
            int numberOfBills;
            while (true)
            {
                Console.Write($"How many new bills do you have for {weekString}? ");
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
            return numberOfBills;
        }

        static string GetBillName(string weekString, int i)
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
            return billName;
        }

        static decimal GetBillAmount(string billName)
        {
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
            return amount;
        }

        static bool GetBillSplitStatus(string billName)
        {
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
            return isSplit;
        }

        static string GetAutopayStatus(string billName)
        {
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
            return autoPayStatus;
        }
    }
}
