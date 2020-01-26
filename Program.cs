using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
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

        private static string date;
        private static List<string[]> callsTable, putsTable;

        static void Main()
        {
            Console.WriteLine("Hello, " + username);

            string optionsURL = base_url + ticker + "/options?p=" + ticker;
            QueryData(optionsURL, true);

            Console.WriteLine("Enter the number next to the date:");
            date = Console.ReadLine();

            string dateURL = optionsURL + "&date=" + date;
            QueryData(dateURL, false);

            ExcelExport(callsTable, putsTable);
        }

        private static void ExcelExport(List<string[]> callsTable, List<string[]> putsTable)
        {
            FileInfo excelFile = new FileInfo(GetFilePath());

            using (ExcelPackage excel = new ExcelPackage(excelFile))
            {
                ExcelWorksheet putsSheet = excel.Workbook.Worksheets["Puts"];
                ExcelWorksheet callsSheet = excel.Workbook.Worksheets["Calls"];

                callsSheet.Cells["AZ3:BJ100"].Value = "";
                putsSheet.Cells["AZ3:BJ100"].Value = "";

                StoreTable(callsSheet, callsTable);
                StoreTable(putsSheet, putsTable);

                excel.Save();
            }
        }

        private static void StoreTable(ExcelWorksheet sheet, List<string[]> table)
        {
            int rowNum = 2;
            string dataRange;
            StringBuilder sb = new StringBuilder();

            foreach (string[] row in table)
            {
                rowNum++;
                dataRange = "AZ" + rowNum + ":BJ" + rowNum;

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

        private static void QueryData(string url, bool printDates)
        {
            HtmlDocument document = new HtmlWeb().Load(url);

            if (printDates)
            {
                DateDropdownParser(document);
            }
            else
            {
                OptionsDataParser(document);
            }
        }

        private static void DateDropdownParser(HtmlDocument document)
        {
            string dateDropdownNode = "//select[@class='Fz(s)']//option";
            
            foreach (HtmlNode node in document.DocumentNode.SelectNodes(dateDropdownNode))
            {
                Console.WriteLine(node.Attributes["value"].Value + " - " + node.InnerText);
            }
        }

        private static void OptionsDataParser(HtmlDocument document)
        {
            string callsTableNode = "//table[@class='calls W(100%) Pos(r) Bd(0) Pt(0) list-options']";
            string putsTableNode = "//table[@class='puts W(100%) Pos(r) list-options']";

            callsTable = document.DocumentNode.SelectSingleNode(callsTableNode)
                .Descendants("tr")
                .Where(tr => tr.Elements("td").Count() > 1)
                .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToArray())
                .ToList();

            putsTable = document.DocumentNode.SelectSingleNode(putsTableNode)
                .Descendants("tr")
                .Where(tr => tr.Elements("td").Count() > 1)
                .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToArray())
                .ToList();
        }

    }
}
