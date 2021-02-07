using System;
using System.Collections.Generic;
using System.IO;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ModifyTaste
{
    class Program
    {
        static void Main(
            string[] args
            )
        {
            //string[] args = new string[2];

            const string quote = "\"";
            const string doubleQuote = quote + quote;
            const string comma = ",";
            const string end = "\r\n";

            // Source: G:\Documents\TetraProject\Database\zerOne_TURN.xlsx
            // Target: G:\Documents\TetraProject\Database\

            if (args is null || args.Length != 2)
            {
                args = new string[2];
            }

            //args[0] = @"G:\Documents\TetraProject\Database\zerOne_TURN.xlsx";
            //args[1] = @"G:\Documents\TetraProject\Database\";

            if (args[0] is null || args[1] is null)
            {
                Console.Write("输入完整的 Database 文件路径：");
                args[0] = Console.ReadLine();
                Console.Write("输入完整的目标文件夹路径（包括末尾的反斜杠）：");
                args[1] = Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Database 文件路径：" + args[0]);
                Console.WriteLine("目标文件夹路径：" + args[1]);
            }

        Start:

            //System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
            //watch.Start();

            FileStream xlsxFileStream = new FileStream(args[0], FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IWorkbook workbook = new XSSFWorkbook(xlsxFileStream);

            Console.WriteLine("\n\n开始转换！");
            Console.WriteLine("\n读取数据库内容：");

            List<List<List<string>>> table = new List<List<List<string>>>();

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                table.Add(new List<List<string>>());// 添加新表
                if (workbook.GetSheetAt(i).SheetName.StartsWith("~"))// 若为临时，跳过
                    continue;
                int handCount = 0;
                bool hand = true;
                for (int j = 1; j < workbook.GetSheetAt(i).LastRowNum + 1; j++)
                {
                    table[i].Add(new List<string>());// 添加新行
                    if (workbook.GetSheetAt(i).GetRow(j) is null)// 若为空，跳过
                        continue;
                    if (hand)// 所有行参照表头
                    {
                        hand = false;
                        handCount = workbook.GetSheetAt(i).GetRow(j).LastCellNum;
                    }
                    for (int k = 0; k < handCount; k++)
                    {
                        if (workbook.GetSheetAt(i).GetRow(j).GetCell(k) is null || workbook.GetSheetAt(i).GetRow(j).GetCell(k).CellType is CellType.Formula)// 若为空公式，写入
                            table[i][j - 1].Add("");
                        else
                            table[i][j - 1].Add(workbook.GetSheetAt(i).GetRow(j).GetCell(k).ToString());// 添加元素
                    }
                }
                Console.WriteLine("    " + workbook.GetSheetAt(i).SheetName + " 已读取；");
            }

            Console.WriteLine("  读取完成。");
            Console.WriteLine("\n将 xlsx 转换为 csv 格式：");

            List<string> csvs = new List<string>();

            for (int i = 0; i < table.Count; i++)// 遍历表
            {
                if (workbook.GetSheetAt(i).SheetName.StartsWith("~"))// 若为临时，跳过
                    continue;
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
                Console.WriteLine("    " + workbook.GetSheetAt(i).SheetName + " 已转换；");
            }

            Console.WriteLine("  转换完成。");
            Console.WriteLine("\n将字符串写入 csv 文件：");

            for (int i = 0; i < csvs.Count; i++)
                File.WriteAllText(args[1] + workbook.GetSheetAt(i).SheetName + ".csv", csvs[i], System.Text.Encoding.UTF8);

            Console.WriteLine("  " + args[0] + " 写入完成。");
            Console.WriteLine("\n成功！");

            //watch.Stop();
            //TimeSpan timeSpan = watch.Elapsed;
            //System.Diagnostics.Debug.WriteLine("代码执行时间：" + timeSpan.TotalMilliseconds);

            Console.Write("\n继续？（Y/n）");
            if (Console.ReadLine().Trim().ToLower() is "n")
                return;
            else
                goto Start;
        }
    }
}
