using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace ConsoleApp1
{
    public static class Utils
    {
        public static IEnumerable<DataAnalysis> GetData(this ExcelPackage package, string sheetName)
        {
            var dataSetSheet = package.Workbook.Worksheets[sheetName];
            var start = dataSetSheet.Dimension.Start;
            var end = dataSetSheet.Dimension.End;

            var data = new List<DataAnalysis>();

            for (int row = start.Row + 1; row <= end.Row; row++)
            {
                var dataRow = new DataAnalysis();
                for (int col = start.Column; col <= end.Column; col++)
                {
                    switch (col)
                    {
                        case 1:
                            dataRow.Sold = dataSetSheet.Cells[row, col].Text;
                            break;
                        case 2:
                            dataRow.Purchased = dataSetSheet.Cells[row, col].Text;
                            break;
                        case 3:
                            dataRow.Value = int.Parse(dataSetSheet.Cells[row, col].Text);
                            break;
                        case 4:
                            dataRow.Owner = dataSetSheet.Cells[row, col].Text;
                            break;
                    }
                }

                data.Add(dataRow);
            }

            return data;
        }
    }
}


public class DataAnalysis
{
    public string Sold { get; set; }
    public string Purchased { get; set; }
    public int Value { get; set; }
    public string Owner { get; set; }

}