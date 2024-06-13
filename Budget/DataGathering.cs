namespace Budget
{
    internal class DataGathering
    {
        public Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>> CollectBills()
        {
            var existingBills = new Dictionary<int, List<(string billName, decimal amount, bool isSplit, string autopayStatus)>>();

            for (int week = 1; week <= 4; week++)
            {
                string weekString = $"Week {week}";
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

                for (int i = 0; i < numberOfBills; i++)
                {
                    string? billName = GetBillName(weekString, i);
                    decimal amount = GetBillAmount(billName);
                    bool isSplit = IsBillSplit(billName);
                    string? autoPayStatus = GetAutoPayStatus(billName);

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

            return existingBills;
        }

        private string GetBillName(string weekString, int index)
        {
            string? billName;
            while (true)
            {
                Console.Write($"Enter the name of bill {index + 1} for {weekString}: ");
                billName = Console.ReadLine();

                if (!string.IsNullOrWhiteSpace(billName))
                {
                    return billName;
                }
                else
                {
                    Console.WriteLine("Invalid input. Bill name cannot be empty.");
                }
            }
        }

        private decimal GetBillAmount(string billName)
        {
            decimal amount;
            while (true)
            {
                Console.Write($"Enter the amount for {billName}: ");
                string? input = Console.ReadLine();

                if (decimal.TryParse(input, out amount) && amount >= 0)
                {
                    return amount;
                }
                else
                {
                    Console.WriteLine("Invalid input. Please enter a non-negative decimal number.");
                }
            }
        }

        private bool IsBillSplit(string billName)
        {
            while (true)
            {
                Console.Write($"Are you splitting {billName} with a roommate? (yes/no): ");
                string? input = Console.ReadLine().Trim().ToLower();

                if (input == "yes")
                {
                    return true;
                }
                else if (input == "no")
                {
                    return false;
                }
                else
                {
                    Console.WriteLine("Invalid input. Please enter 'yes' or 'no'.");
                }
            }
        }

        private string GetAutoPayStatus(string billName)
        {
            while (true)
            {
                Console.Write($"Enter autopay status for {billName} (yes/no): ");
                string? input = Console.ReadLine().Trim().ToLower();

                if (input == "yes" || input == "no")
                {
                    return input;
                }
                else
                {
                    Console.WriteLine("Invalid input. Please enter 'yes' or 'no'.");
                }
            }
        }
    }

}
