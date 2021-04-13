﻿using ClosedXML.Excel;
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
                $"16 To 49\t" +
                $"50 To 54\t" +
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
                log.Info($"{area.Code}\t{area.Name}\t" +
                    $"{area.Population16To49 / pop.Population16To49:P2}\t" +
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

            var ltla = GetLtlaVaccinations(xb);
            foreach (var area in ltla)
            {
                var pop = populationEstimates[area.Code];
                log.Info($"{area.Code}\t{area.Name}\t" +
                    $"{area.Population16To49 / pop.Population16To49:P2}\t" +
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

            var path = doc.DocumentNode.SelectSingleNode("//h3[text() = 'Weekly data:']/following-sibling::p/a").Attributes["href"].Value;

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

        private static IEnumerable<StatisticalArea> GetLtlaPopulationEstimates(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("Population estimates (NIMS)", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            // Verify headings are as we expect.
            const string lastColumn = "L";
            var sanityCheck = xs.Cell($"{lastColumn}13");
            if (!string.Equals(sanityCheck.Value.ToString(), "80+", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sheet not in the expected format - have additional age bands been added?");
            }

            var range = xs.Range($"B16", $"{lastColumn}329");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To49 = (double)row.Cell(4).Value,
                    Population50To54 = (double)row.Cell(5).Value,
                    Population55To59 = (double)row.Cell(6).Value,
                    Population60To64 = (double)row.Cell(7).Value,
                    Population65To69 = (double)row.Cell(8).Value,
                    Population70To74 = (double)row.Cell(9).Value,
                    Population75To79 = (double)row.Cell(10).Value,
                    PopulationOver80 = (double)row.Cell(11).Value
                };

                yield return result;
            }
        }

        private static IEnumerable<StatisticalArea> GetLtlaVaccinations(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("LTLA", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            // Verify headings are as we expect.
            const string lastColumn = "M";
            var sanityCheck = xs.Cell($"{lastColumn}13");
            if (!string.Equals(sanityCheck.Value.ToString(), "80+", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sheet not in the expected format - have additional age bands been added?");
            }

            var range = xs.Range("D16", $"{lastColumn}329");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To49 = (double)row.Cell(3).Value,
                    Population50To54 = (double)row.Cell(4).Value,
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

        private static IEnumerable<StatisticalArea> GetMsoaPopulationEstimates(XLWorkbook xb)
        {
            if (!xb.TryGetWorksheet("Population estimates (NIMS)", out IXLWorksheet xs))
            {
                throw new InvalidOperationException("Really were expected that tab to exist. Boo.");
            }

            // Verify headings are as we expect.
            const string startColumn = "O";
            var sanityCheck = xs.Cell($"{startColumn}10");
            if (!string.Equals(sanityCheck.Value.ToString(), "NIMS population mapped to MSOA", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sheet not in the expected format - have additional age bands been added?");
            }

            var range = xs.Range($"{startColumn}16", "Y6806");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To49 = (double)row.Cell(4).Value,
                    Population50To54 = (double)row.Cell(5).Value,
                    Population55To59 = (double)row.Cell(6).Value,
                    Population60To64 = (double)row.Cell(7).Value,
                    Population65To69 = (double)row.Cell(8).Value,
                    Population70To74 = (double)row.Cell(9).Value,
                    Population75To79 = (double)row.Cell(10).Value,
                    PopulationOver80 = (double)row.Cell(11).Value
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

            // Verify headings are as we expect.
            const string lastColumn = "O";
            var sanityCheck = xs.Cell($"{lastColumn}13");
            if (!string.Equals(sanityCheck.Value.ToString(), "80+", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Excel sheet not in the expected format - have additional age bands been added?");
            }

            var range = xs.Range("F16", $"{lastColumn}6806");
            var rowCount = range.RowCount();

            for (var i = 1; i <= rowCount; i++)
            {
                var row = range.Row(i);

                var result = new StatisticalArea
                {
                    Code = row.Cell(1).Value.ToString(),
                    Name = row.Cell(2).Value.ToString(),

                    Population16To49 = (double)row.Cell(3).Value,
                    Population50To54 = (double)row.Cell(4).Value,
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
    }
}