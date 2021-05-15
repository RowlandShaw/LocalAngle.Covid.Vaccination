using ClosedXML.Excel;
using HtmlAgilityPack;
using LocalAngle.Covid.Vaccination.Models;
using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;

namespace LocalAngle.Covid.Vaccination.Agent
{
    internal static class Program
    {
        private const int LastLtlaRow = 329;
        private const int LastMsoaRow = 6806;
        private static readonly Uri FileList = new Uri("https://www.england.nhs.uk/statistics/statistical-work-areas/covid-19-vaccinations/");
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));

        private static void Main()
        {
            XmlConfigurator.Configure();

            log.Debug("Starting");

            var mostRecentSpreadsheetUrl = GetMostRecentSpreadsheetUrl();
            var xb = LoadExcel(mostRecentSpreadsheetUrl);
            var populationEstimates = new Dictionary<string, StatisticalArea>();
            var pops = GetMsoaPopulationEstimates(xb);
            foreach (var est in pops)
            {
                populationEstimates.Add(est.Code, est);
            }
            pops = GetLtlaPopulationEstimates(xb);
            foreach (var est in pops)
            {
                populationEstimates.Add(est.Code, est);
            }

            log.Info($"Code\tName\t" +
                $"16 To 39\t" +
                $"40 To 44\t" +
                $"45 To 49\t" +
                $"50 To 54\t" +
                $"55 To 59\t" +
                $"60 To 64\t" +
                $"65 To 69\t" +
                $"70 To 74\t" +
                $"75 To 79\t" +
                $"Over 80\tOverall");

            foreach (var area in GetMsoaFirstVaccinations(xb))
            {
                var pop = populationEstimates[area.Code];
                log.Info($"{area.Code}\t{area.Name}\t" +
                    $"{area.Population16To39 / pop.Population16To39:P2}\t" +
                    $"{area.Population40To44 / pop.Population40To44:P2}\t" +
                    $"{area.Population45To49 / pop.Population45To49:P2}\t" +
                    $"{area.Population50To54 / pop.Population50To54:P2}\t" +
                    $"{area.Population55To59 / pop.Population55To59:P2}\t" +
                    $"{area.Population60To64 / pop.Population60To64:P2}\t" +
                    $"{area.Population65To69 / pop.Population65To69:P2}\t" +
                    $"{area.Population70To74 / pop.Population70To74:P2}\t" +
                    $"{area.Population75To79 / pop.Population75To79:P2}\t" +
                    $"{area.PopulationOver80 / pop.PopulationOver80:P2}\t" +
                    $"{area.PopulationOverall / pop.PopulationOverall:P2}"
                    );
            }

            foreach (var area in GetLtlaFirstVaccinations(xb))
            {
                var pop = populationEstimates[area.Code];
                log.Info($"{area.Code}\t{area.Name}\t" +
                    $"{area.Population16To39 / pop.Population16To39:P2}\t" +
                    $"{area.Population40To44 / pop.Population40To44:P2}\t" +
                    $"{area.Population45To49 / pop.Population45To49:P2}\t" +
                    $"{area.Population50To54 / pop.Population50To54:P2}\t" +
                    $"{area.Population55To59 / pop.Population55To59:P2}\t" +
                    $"{area.Population60To64 / pop.Population60To64:P2}\t" +
                    $"{area.Population65To69 / pop.Population65To69:P2}\t" +
                    $"{area.Population70To74 / pop.Population70To74:P2}\t" +
                    $"{area.Population75To79 / pop.Population75To79:P2}\t" +
                    $"{area.PopulationOver80 / pop.PopulationOver80:P2}\t" +
                    $"{area.PopulationOverall / pop.PopulationOverall:P2}"
                    );
            }

            foreach (var area in GetLtlaSecondVaccinations(xb))
            {
                var pop = populationEstimates[area.Code];
                log.Info($"{area.Code}\t{area.Name}\t" +
                    $"{area.Population16To39 / pop.Population16To39:P2}\t" +
                    $"{area.Population40To44 / pop.Population40To44:P2}\t" +
                    $"{area.Population45To49 / pop.Population45To49:P2}\t" +
                    $"{area.Population50To54 / pop.Population50To54:P2}\t" +
                    $"{area.Population55To59 / pop.Population55To59:P2}\t" +
                    $"{area.Population60To64 / pop.Population60To64:P2}\t" +
                    $"{area.Population65To69 / pop.Population65To69:P2}\t" +
                    $"{area.Population70To74 / pop.Population70To74:P2}\t" +
                    $"{area.Population75To79 / pop.Population75To79:P2}\t" +
                    $"{area.PopulationOver80 / pop.PopulationOver80:P2}\t" +
                    $"{area.PopulationOverall / pop.PopulationOverall:P2}"
                    );
            }

            log.Debug("Complete");
        }

        private static Uri GetMostRecentSpreadsheetUrl()
        {
            log.Debug($"Determining latest data available from {FileList}");

            var client = new HtmlWeb();
            var doc = client.Load(FileList);

            var path = doc.DocumentNode.SelectSingleNode("//p[strong/text() = 'Weekly data']/following-sibling::p/a").Attributes["href"].Value;

            return new Uri(FileList, path);
        }

        private static XLWorkbook LoadExcel(Uri uri)
        {
            log.Debug($"Downloading {uri}");
            var client = new HttpClient();

            var response = client.GetAsync(uri).Result;
            var str = response.Content.ReadAsStreamAsync().Result;
            return LoadExcel(str);
        }

        private static XLWorkbook LoadExcel(Stream str)
        {
            return new XLWorkbook(str);
        }

        private static IEnumerable<StatisticalArea> GetFirstVaccinations(IXLWorksheet xs, int lastRow)
        {
            // Verify headings are as we expect.
            const string lastColumn = "Q";
            var sanityCheck = xs.Cell($"{lastColumn}13");
            if (!string.Equals(sanityCheck.Value.ToString(), "80+", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sheet not in the expected format - have additional age bands been added?");
            }

            var range = xs.Range("F16", $"{lastColumn}{lastRow}");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To39 = (double)row.Cell(3).Value,
                    Population40To44 = (double)row.Cell(4).Value,
                    Population45To49 = (double)row.Cell(5).Value,
                    Population50To54 = (double)row.Cell(6).Value,
                    Population55To59 = (double)row.Cell(7).Value,
                    Population60To64 = (double)row.Cell(8).Value,
                    Population65To69 = (double)row.Cell(9).Value,
                    Population70To74 = (double)row.Cell(10).Value,
                    Population75To79 = (double)row.Cell(11).Value,
                    PopulationOver80 = (double)row.Cell(12).Value
                };

                yield return result;
            }
        }

        private static IEnumerable<StatisticalArea> GetLtlaPopulationEstimates(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("Population estimates (NIMS)", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            // Verify headings are as we expect.
            const string lastColumn = "P";
            var sanityCheck = xs.Cell($"{lastColumn}13");
            if (!string.Equals(sanityCheck.Value.ToString(), "80+", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sheet not in the expected format - have additional age bands been added?");
            }

            var range = xs.Range($"D16", $"{lastColumn}{LastLtlaRow}");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To39 = (double)row.Cell(4).Value,
                    Population40To44 = (double)row.Cell(5).Value,
                    Population45To49 = (double)row.Cell(6).Value,
                    Population50To54 = (double)row.Cell(7).Value,
                    Population55To59 = (double)row.Cell(8).Value,
                    Population60To64 = (double)row.Cell(9).Value,
                    Population65To69 = (double)row.Cell(10).Value,
                    Population70To74 = (double)row.Cell(11).Value,
                    Population75To79 = (double)row.Cell(12).Value,
                    PopulationOver80 = (double)row.Cell(13).Value
                };

                yield return result;
            }
        }

        private static IEnumerable<StatisticalArea> GetLtlaFirstVaccinations(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("LTLA", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            foreach (var v in GetFirstVaccinations(xs, LastLtlaRow)) { yield return v; }
        }

        private static IEnumerable<StatisticalArea> GetLtlaSecondVaccinations(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("LTLA", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            foreach (var v in GetSecondVaccinations(xs, LastLtlaRow)) { yield return v; }
        }

        private static IEnumerable<StatisticalArea> GetMsoaPopulationEstimates(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("Population estimates (NIMS)", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            // Verify headings are as we expect.
            const string startColumn = "S";
            var sanityCheck = xs.Cell($"{startColumn}10");
            if (!string.Equals(sanityCheck.Value.ToString(), "NIMS population mapped to MSOA", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sheet not in the expected format - have additional age bands been added?");
            }

            var range = xs.Range($"{startColumn}16", $"AF{LastMsoaRow}");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To39 = (double)row.Cell(4).Value,
                    Population40To44 = (double)row.Cell(5).Value,
                    Population45To49 = (double)row.Cell(6).Value,
                    Population50To54 = (double)row.Cell(7).Value,
                    Population55To59 = (double)row.Cell(8).Value,
                    Population60To64 = (double)row.Cell(9).Value,
                    Population65To69 = (double)row.Cell(10).Value,
                    Population70To74 = (double)row.Cell(11).Value,
                    Population75To79 = (double)row.Cell(12).Value,
                    PopulationOver80 = (double)row.Cell(13).Value
                };

                yield return result;
            }
        }

        private static IEnumerable<StatisticalArea> GetMsoaFirstVaccinations(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("MSOA", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            foreach (var v in GetFirstVaccinations(xs, LastMsoaRow)) { yield return v; }
        }

        private static IEnumerable<StatisticalArea> GetSecondVaccinations(IXLWorksheet xs, int lastRow)
        {
            // Verify headings are as we expect.
            const string lastColumn = "AB";
            var sanityCheck = xs.Cell($"{lastColumn}13");
            if (!string.Equals(sanityCheck.Value.ToString(), "80+", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sheet not in the expected format - have additional age bands been added?");
            }

            var range = xs.Range("F16", $"{lastColumn}{lastRow}");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To39 = (double)row.Cell(14).Value,
                    Population40To44 = (double)row.Cell(15).Value,
                    Population45To49 = (double)row.Cell(16).Value,
                    Population50To54 = (double)row.Cell(17).Value,
                    Population55To59 = (double)row.Cell(18).Value,
                    Population60To64 = (double)row.Cell(19).Value,
                    Population65To69 = (double)row.Cell(20).Value,
                    Population70To74 = (double)row.Cell(21).Value,
                    Population75To79 = (double)row.Cell(22).Value,
                    PopulationOver80 = (double)row.Cell(23).Value
                };

                yield return result;
            }
        }
    }
}