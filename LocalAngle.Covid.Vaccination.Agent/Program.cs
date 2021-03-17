using ClosedXML.Excel;
using HtmlAgilityPack;
using LocalAngle.Covid.Vaccination.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;

namespace LocalAngle.Covid.Vaccination.Agent
{
    internal static class Program
    {
        private static readonly Uri FileList = new Uri("https://www.england.nhs.uk/statistics/statistical-work-areas/covid-19-vaccinations/");

        private static void Main()
        {
            Console.WriteLine("Hello World!");
            var mostRecentSpreadsheetUrl = GetMostRecentSpreadsheetUrl();
            var xb = LoadExcel(mostRecentSpreadsheetUrl);
            var msoa = GetMsoaVaccinations(xb);
            foreach (var area in msoa)
            {
                Console.WriteLine($"{area.Code} ({area.Name}) {area.Population16To59}");
            }
            Console.ReadKey();
        }

        private static Uri GetMostRecentSpreadsheetUrl()
        {
            var client = new HtmlWeb();
            var doc = client.Load(FileList);

            var path = doc.DocumentNode.SelectSingleNode("//h3[text() = 'Weekly data:']/following-sibling::p/a").Attributes["href"].Value;

            return new Uri(FileList, path);
        }

        private static XLWorkbook LoadExcel(Uri uri)
        {
            var client = new HttpClient();

            var response = client.GetAsync(uri).Result;
            var str = response.Content.ReadAsStreamAsync().Result;
            return LoadExcel(str);
        }

        private static XLWorkbook LoadExcel(Stream str)
        {
            return new XLWorkbook(str);
        }

        private static IEnumerable<StatisticalArea> GetMsoaVaccinations(XLWorkbook xb)
        {
            IXLWorksheet xs;
            if (!xb.TryGetWorksheet("MSOA", out xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            // Data should start be in F16:M6806
            var range = xs.Range("F16", "M6806");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea();

                result.Code = row.Cell(1).Value.ToString();
                result.Name = row.Cell(2).Value.ToString();

                result.Population16To59 = Convert.ToInt32(row.Cell(3).Value);
                result.Population60To64 = Convert.ToInt32(row.Cell(4).Value);
                result.Population65To69 = Convert.ToInt32(row.Cell(5).Value);
                result.Population70To74 = Convert.ToInt32(row.Cell(6).Value);
                result.Population75To79 = Convert.ToInt32(row.Cell(7).Value);
                result.PopulationOver80 = Convert.ToInt32(row.Cell(8).Value);

                yield return result;
            }
        }
    }
}