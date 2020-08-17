using System;
using System.Collections.Generic;
using System.IO;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace TetraModifyTaste
{
    class Program
    {
        static void Main(
            string[] args
            )
        {
            //string[] args = new string[2];
            //args[0] = @"G:\Project\TetraModifyTaste\Database.xlsx";
            //args[1] = @"G:\Project\TetraModifyTaste\";

            const string quote = "\"";
            const string doubleQuote = quote + quote;
            const string comma = ",";
            const string end = "\r\n";

            FileInfo xlsxFile = new FileInfo(args[0]);

        begin:

            IWorkbook workbook = new XSSFWorkbook(xlsxFile);
            List<List<List<string>>> table = new List<List<List<string>>>();

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                table.Add(new List<List<string>>());// 添加新表
                int handCount = 0;
                bool hand = true;
                for (int j = 0; j < workbook.GetSheetAt(i).LastRowNum; j++)
                {
                    table[i].Add(new List<string>());// 添加新行
                    if (workbook.GetSheetAt(i).GetRow(j) is null) //若为空，跳过
                        continue;
                    if (hand)// 所有行参照表头
                    {
                        hand = false;
                        handCount = workbook.GetSheetAt(i).GetRow(j).LastCellNum;
                    }
                    for (int k = 0; k < handCount; k++)
                    {
                        if (workbook.GetSheetAt(i).GetRow(j).GetCell(k) is null)// 若为空，写入
                            table[i][j].Add("");
                        else
                            table[i][j].Add(workbook.GetSheetAt(i).GetRow(j).GetCell(k).ToString());// 添加元素
                    }
                }
            }

            List<string> csvs = new List<string>();

            for (int i = 0; i < table.Count; i++)// 遍历表
            {
                string csv = "";
                for (int j = 0; j < table[i].Count; j++)// 遍历行
                {
                    string row = "";
                    bool first = true;
                    for (int k = 0; k < table[i][j].Count; k++)// 遍历元素
                    {
                        string cell = table[i][j][k];
                        if (cell.Contains(quote) | cell.Contains(comma) | cell.Contains("\n"))// 若存在引号、逗号或换行符
                            cell = quote + cell.Replace(quote, doubleQuote) + quote;
                        if (first)
                            first = false;
                        else
                            cell = comma + cell;
                        row += cell;
                    }
                    if (!(row is ""))
                        csv += row + end;
                }
                csvs.Add(csv);
            }

            for (int i = 0; i < csvs.Count; i++)
                File.WriteAllText(args[1] + workbook.GetSheetAt(i).SheetName + ".csv", csvs[i], System.Text.Encoding.UTF8);

            Console.WriteLine("Reset? (y/N) ");
            if (Console.ReadKey().Key is ConsoleKey.Y)
                goto begin;
        }
    }
}