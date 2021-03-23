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
            var mostRecentSpreadsheetUrl = GetMostRecentSpreadsheetUrl();
            var xb = LoadExcel(mostRecentSpreadsheetUrl);
            var pops = GetPopulationEstimates(xb);
            var populationEstimates = new Dictionary<string, StatisticalArea>();
            foreach (var est in pops)
            {
                populationEstimates.Add(est.Code, est);
            }

            Console.WriteLine($"Code\tName\t" +
                $"16 To 54\t" +
                $"55 To 59\t" +
                $"60 To 64\t" +
                $"65 To 69\t" +
                $"70 To 74\t" +
                $"75 To 79\t" +
                $"Over 80\tOverall");
            var msoa = GetMsoaVaccinations(xb);
            foreach (var area in msoa)
            {
                var pop = populationEstimates[area.Code];
                Console.WriteLine($"{area.Code}\t{area.Name}\t" +
                    $"{area.Population16To54 / pop.Population16To54:P2}\t" +
                    $"{area.Population55To59 / pop.Population55To59:P2}\t" +
                    $"{area.Population60To64 / pop.Population60To64:P2}\t" +
                    $"{area.Population65To69 / pop.Population65To69:P2}\t" +
                    $"{area.Population70To74 / pop.Population70To74:P2}\t" +
                    $"{area.Population75To79 / pop.Population75To79:P2}\t" +
                    $"{area.PopulationOver80 / pop.PopulationOver80:P2}\t" +
                    $"{area.PopulationOverall / pop.PopulationOverall:P2}"
                    );
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

        private static IEnumerable<StatisticalArea> GetPopulationEstimates(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("Population estimates (NIMS)", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            // Data should start be in F16:M6806
            var range = xs.Range("N16", "W6806");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To54 = (double)row.Cell(4).Value,
                    Population55To59 = (double)row.Cell(5).Value,
                    Population60To64 = (double)row.Cell(6).Value,
                    Population65To69 = (double)row.Cell(7).Value,
                    Population70To74 = (double)row.Cell(8).Value,
                    Population75To79 = (double)row.Cell(9).Value,
                    PopulationOver80 = (double)row.Cell(10).Value
                };

                yield return result;
            }
        }

        private static IEnumerable<StatisticalArea> GetMsoaVaccinations(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("MSOA", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            // Data should start be in F16:M6806
            var range = xs.Range("F16", "N6806");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To54 = (double)row.Cell(3).Value,
                    Population55To59 = (double)row.Cell(4).Value,
                    Population60To64 = (double)row.Cell(5).Value,
                    Population65To69 = (double)row.Cell(6).Value,
                    Population70To74 = (double)row.Cell(7).Value,
                    Population75To79 = (double)row.Cell(8).Value,
                    PopulationOver80 = (double)row.Cell(9).Value
                };

                yield return result;
            }
        }
    }
}