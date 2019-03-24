using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Metadata.Ecma335;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            List<DataAnalysis> data;

            var appRoot = AppRoot();
            var filePath = Path.Combine(appRoot, "dataset.xlsx");
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                data = package.GetData("dataset").ToList();
            }

            GetDistinctCombinations(data);
            Transactions500(data);
            ThirdHiestAverage(data);
        }

        private static void ThirdHiestAverage(List<DataAnalysis> data)
        {
            var grouping = data
                .GroupBy(x => x.Owner)
                .ToDictionary(k => k.Key, v => v.Select(x => x.Purchased));

            var accounts = new Dictionary<string, int>();
            foreach (var group in grouping)
            {
                var distinct = group.Value.Distinct().Count();
                var fullCount = group.Value.Count();
                var avg = fullCount / distinct;

                accounts.Add(group.Key, avg);
            }

            var thirdAvrg = accounts.OrderBy(x => x.Value).Skip(2).Take(1).FirstOrDefault();
            Console.WriteLine($"3. Third Average account: {thirdAvrg.Key}");
        }

        private static void Transactions500(List<DataAnalysis> data)
        {
            var accounts = data.GroupBy(x => x.Owner).Where(x => x.Count() >= 500).ToList();

            var allaccounts = string.Join(",", accounts.Select(x=>x.Key));
            Console.WriteLine($"2. Accounts with at least 500 transactions: {allaccounts}");
        }

        private static void GetDistinctCombinations(List<DataAnalysis> data)
        {
            var grouping = data
                .GroupBy(x => x.Sold)
                .ToDictionary(k => k.Key, v => v.Select(x => x.Purchased));
            int total = 0;
            foreach (var group in grouping)
            {
                total += group.Value.Distinct().Count();
            }
            Console.WriteLine($"1. Differet Combinatins: {total}");
        }

        private static string AppRoot()
        {
            var exePath = Path.GetDirectoryName(System.Reflection
                .Assembly.GetExecutingAssembly().CodeBase);
            Regex appPathMatcher = new Regex(@"(?<!fil)[A-Za-z]:\\+[\S\s]*?(?=\\+bin)");
            var appRoot = appPathMatcher.Match(exePath).Value;
            return appRoot;
        }
    }
}


