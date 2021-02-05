using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace OptionsUpdater
{
    class Program
    {
        private static readonly string username = "dylan";
        private static readonly string base_url = "https://finance.yahoo.com/quote/";
        private static readonly string ticker = "UVXY";

        private static readonly string currentPriceNode = "//span[@class='Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)']";
        private static readonly string dateDropdownNode = "//select[@class='Fz(s) H(25px) Bd Bdc($seperatorColor)']//option";
        private static readonly string callsTableNode = "//table[@class='calls W(100%) Pos(r) Bd(0) Pt(0) list-options']";
        //private static readonly string putsTableNode = "//table[@class='puts W(100%) Pos(r) list-options']";

        private static Dictionary<int, int> unixTimestamp = new Dictionary<int, int>();
        private static Dictionary<int, string> dateFormat = new Dictionary<int, string>();
        private static List<string[]> callsTable, callsTable2;

        private static string currentPrice;

        static void Main()
        {
            Console.WriteLine("Hello, " + username);
            
            string optionsURL = base_url + ticker + "/options?p=" + ticker;
            QueryDates(optionsURL);

            Console.WriteLine("Enter the number next to the date:");
            int choice = int.Parse(Console.ReadLine());

            string dateURL = optionsURL + "&date=" + unixTimestamp[choice];
            callsTable = QueryData(dateURL, callsTableNode);

            Console.WriteLine("Enter the number next to the date:");
            int choice2 = int.Parse(Console.ReadLine());

            dateURL = optionsURL + "&date=" + unixTimestamp[choice2];
            callsTable2 = QueryData(dateURL, callsTableNode);


            FileInfo excelFile = new FileInfo(GetFilePath());

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage(excelFile))
            {
                ExcelExport(excel, callsTable, "Calls", choice);
                ExcelExport(excel, callsTable2, "Calls2", choice2);
                excel.Save();
            }

            Process psExcel = new Process();
            psExcel.StartInfo.UseShellExecute = true;
            psExcel.StartInfo.FileName = GetFilePath();
            psExcel.Start();

            Process.GetCurrentProcess().Kill();
        }

        private static void ExcelExport(ExcelPackage excel, List<string[]> table, string sheetName, int choice)
        {
            ExcelWorksheet sheet = excel.Workbook.Worksheets[sheetName];

            sheet.Cells["A1:L1"].Clear();

            sheet.Cells["A1"].Value = "Calls for " + dateFormat[choice];
            sheet.Cells["B1"].Value = ticker;
            sheet.Cells["C1"].Value = decimal.Round(decimal.Parse(currentPrice), 2);
            sheet.Cells["L1"].Value = dateFormat[choice];

            sheet.Cells["A3:K100"].Value = null;

            StoreTable(sheet, table);
        }

        private static void StoreTable(ExcelWorksheet sheet, List<string[]> table)
        {
            int rowNum = 2;
            string dataRange;
            StringBuilder sb = new StringBuilder();

            foreach (string[] row in table)
            {
                rowNum++;
                dataRange = "A" + rowNum + ":K" + rowNum;

                foreach (string cell in row)
                {
                    sb.Append(cell.Replace(",", "") + ",");
                }

                sheet.Cells[dataRange].LoadFromText(sb.ToString());
                sb.Clear();
            }
        }

        private static string GetFilePath()
        {
            return @"C:\Users\" + username + @"\Desktop\test.xlsx";
        }

        private static void QueryDates(string url)
        {
            HtmlDocument document = new HtmlWeb().Load(url);
            CurrentPriceParser(document);
            DateDropdownParser(document);
        }

        private static List<string[]> QueryData(string url, string node)
        {
            HtmlDocument document = new HtmlWeb().Load(url);

            return OptionsDataParser(document, node);
        }

        private static void CurrentPriceParser(HtmlDocument document)
        {
            HtmlNode priceNode = document.DocumentNode.SelectSingleNode(currentPriceNode);
            currentPrice = priceNode.InnerText;

            Console.WriteLine("The current price for " + ticker + " is: $" + currentPrice);
        }

        private static void DateDropdownParser(HtmlDocument document)
        {
            int i = 0;

            foreach (HtmlNode node in document.DocumentNode.SelectNodes(dateDropdownNode))
            {
                i++;
                Console.WriteLine("(" + i + ") - " + node.InnerText);

                unixTimestamp.Add(i, int.Parse(node.Attributes["value"].Value));
                dateFormat.Add(i, node.InnerText);
            }
        }

        private static List<string[]> OptionsDataParser(HtmlDocument document, string tableNode)
        {
            return document.DocumentNode.SelectSingleNode(tableNode)
                .Descendants("tr")
                .Where(tr => tr.Elements("td").Count() > 1)
                .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToArray())
                .ToList();
        }

    }
}
