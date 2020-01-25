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
        private static readonly string username = "Gatsby";
        private static readonly string base_url = "https://finance.yahoo.com/quote/";
        private static readonly string ticker = "UVXY";

        private static List<string[]> headerRow = new List<string[]>(){
            new string[] { "Contract Name","Last Trade Date","Strike","Last Price","Bid","Ask","Change  %","Change","Volume","Open Interest","Implied Volatility" }
        };

        private static string date;
        private static List<List<string>> callsTable, putsTable;

        static void Main(string[] args)
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

        private static void ExcelExport(List<List<string>> callsTable, List<List<string>> putsTable)
        {
            FileInfo excelFile = new FileInfo(GetFilePath());

            using (ExcelPackage excel = new ExcelPackage(excelFile))
            {
                ExcelWorksheet putsSheet = excel.Workbook.Worksheets["Puts"];
                ExcelWorksheet callsSheet = excel.Workbook.Worksheets["Calls"];

                callsSheet.Cells["AZ2:BJ140"].Value = "";
                putsSheet.Cells["AZ2:BJ140"].Value = "";

                string headerRange = "AZ2:BJ2";
                callsSheet.Cells[headerRange].LoadFromArrays(headerRow);
                putsSheet.Cells[headerRange].LoadFromArrays(headerRow);

                int rowNum = 2;
                string dataRange;
                foreach (List<string> row in callsTable)
                {
                    rowNum++;
                    dataRange = "AZ" + rowNum + ":" + "BJ" + rowNum;
                    StringBuilder sb = new StringBuilder();
                    foreach (string cell in row)
                    {
                        sb.Append(cell.Replace(",", "").Trim() + ",");
                    }
                    callsSheet.Cells[dataRange].LoadFromText(sb.ToString());
                }

                rowNum = 2;
                foreach (List<string> row in putsTable)
                {
                    rowNum++;
                    dataRange = "AZ" + rowNum + ":" + "BJ" + rowNum;
                    StringBuilder sb = new StringBuilder();
                    foreach (string cell in row)
                    {
                        sb.Append(cell.Replace(",", "").Trim() + ",");
                    }
                    putsSheet.Cells[dataRange].LoadFromText(sb.ToString());
                }

                excel.Save();
            }
        }

        private static string GetFilePath()
        {
            return @"C:\Users\" + username + @"\Desktop\UVXY Put Option Payoff.xlsx";
        }


        private static void QueryData(string url, bool printDates)
        {
            HtmlDocument document = new HtmlWeb().Load(url);
            OptionsDataParser(document, printDates);
        }

        private static void OptionsDataParser(HtmlDocument document, bool printDates)
        {
            string dateDropdownNode = "//select[@class='Fz(s)']//option";

            string callsTableNode = "//table[@class='calls W(100%) Pos(r) Bd(0) Pt(0) list-options']";
            string putsTableNode = "//table[@class='puts W(100%) Pos(r) list-options']";

            if (printDates)
            {
                foreach (HtmlNode node in document.DocumentNode.SelectNodes(dateDropdownNode))
                {
                    Console.WriteLine(node.Attributes["value"].Value + " - " + node.InnerText);
                }
            }


            callsTable = document.DocumentNode.SelectSingleNode(callsTableNode)
                .Descendants("tr")
                .Skip(1)
                .Where(tr => tr.Elements("td").Count() > 1)
                .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                .ToList();

            putsTable = document.DocumentNode.SelectSingleNode(putsTableNode)
                .Descendants("tr")
                .Skip(1)
                .Where(tr => tr.Elements("td").Count() > 1)
                .Select(tr => tr.Elements("td").Select(td => td.InnerText.Trim()).ToList())
                .ToList();
        }

    }
}
