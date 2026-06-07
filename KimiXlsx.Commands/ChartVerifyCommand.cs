using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;

namespace KimiXlsx.Commands;

public static class ChartVerifyCommand
{
	private static readonly XNamespace ChartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

	private static readonly Dictionary<string, string> ChartTypes = new Dictionary<string, string>
	{
		{ "barChart", "bar" },
		{ "lineChart", "line" },
		{ "pieChart", "pie" },
		{ "areaChart", "area" },
		{ "scatterChart", "scatter" },
		{ "doughnutChart", "doughnut" },
		{ "radarChart", "radar" }
	};

	public static int Run(string[] args)
	{
		if (args.Length < 1)
		{
			Console.WriteLine("Usage: KimiXlsx chart-verify <excel_file> [--json] [--verbose/-v] [--debug]");
			return 1;
		}
		string filePath = args[0];
		bool num = args.Contains("--json");
		bool verbose = args.Contains("--verbose") || args.Contains("-v");
		bool debug = args.Contains("--debug");
		VerificationResult verificationResult = VerifyChartData(filePath, debug);
		if (num)
		{
			JsonSerializerOptions options = new JsonSerializerOptions
			{
				WriteIndented = true
			};
			Console.WriteLine(JsonSerializer.Serialize(verificationResult, options));
		}
		else
		{
			PrintReport(verificationResult, verbose);
		}
		return (!verificationResult.Valid) ? 1 : 0;
	}

	public static VerificationResult VerifyChartData(string filePath, bool debug = false)
	{
		VerificationResult verificationResult = new VerificationResult
		{
			File = filePath
		};
		try
		{
			using ZipArchive zipArchive = ZipFile.OpenRead(filePath);
			if (debug)
			{
				Console.WriteLine("[DEBUG] All entries in archive:");
				foreach (ZipArchiveEntry item in zipArchive.Entries.Take(50))
				{
					Console.WriteLine("  - " + item.FullName);
				}
				if (zipArchive.Entries.Count > 50)
				{
					Console.WriteLine($"  ... and {zipArchive.Entries.Count - 50} more");
				}
			}
			List<ZipArchiveEntry> list = zipArchive.Entries.Where(delegate(ZipArchiveEntry e)
			{
				string text = e.FullName.Replace('\\', '/').ToLowerInvariant();
				string fileName = Path.GetFileName(text);
				return text.Contains("/charts/") && fileName.StartsWith("chart") && fileName.EndsWith(".xml") && !fileName.Contains("colors") && !fileName.Contains("style");
			}).ToList();
			if (debug)
			{
				Console.WriteLine($"[DEBUG] Found {list.Count} chart files:");
				foreach (ZipArchiveEntry item2 in list)
				{
					Console.WriteLine("  - " + item2.FullName);
				}
			}
			verificationResult.TotalCharts = list.Count;
			if (list.Count == 0)
			{
				verificationResult.Errors.Add("No charts found in workbook");
				return verificationResult;
			}
			foreach (ZipArchiveEntry item3 in list)
			{
				ChartResult chartResult = VerifySingleChart(zipArchive, item3);
				if (chartResult.HasData)
				{
					if (chartResult.HasMeaningfulData)
					{
						verificationResult.ValidCharts.Add(chartResult);
						continue;
					}
					verificationResult.ChartsWithNoData.Add(chartResult);
					verificationResult.Valid = false;
				}
				else
				{
					verificationResult.EmptyCharts.Add(chartResult);
					verificationResult.Valid = false;
				}
			}
		}
		catch (InvalidDataException)
		{
			verificationResult.Valid = false;
			verificationResult.Errors.Add("File is not a valid ZIP/XLSX archive");
		}
		catch (Exception ex2)
		{
			verificationResult.Valid = false;
			verificationResult.Errors.Add("Error reading file: " + ex2.Message);
		}
		return verificationResult;
	}

	private static ChartResult VerifySingleChart(ZipArchive archive, ZipArchiveEntry chartEntry)
	{
		string fileName = Path.GetFileName(chartEntry.FullName);
		ChartResult chartResult = new ChartResult
		{
			File = chartEntry.FullName,
			Name = fileName
		};
		try
		{
			XElement xElement = XDocument.Load(chartEntry.Open()).Descendants(ChartNs + "chart").FirstOrDefault();
			if (xElement == null)
			{
				chartResult.Issues.Add("No <c:chart> element found");
				return chartResult;
			}
			XElement xElement2 = xElement.Descendants(ChartNs + "plotArea").FirstOrDefault();
			if (xElement2 == null)
			{
				chartResult.Issues.Add("No <c:plotArea> element found");
				return chartResult;
			}
			foreach (KeyValuePair<string, string> chartType in ChartTypes)
			{
				chartType.Deconstruct(out var key, out var value);
				string text = key;
				string text2 = value;
				XElement xElement3 = xElement2.Descendants(ChartNs + text).FirstOrDefault();
				if (xElement3 == null)
				{
					continue;
				}
				chartResult.ChartType = text2;
				List<XElement> list = xElement3.Descendants(ChartNs + "ser").ToList();
				chartResult.SeriesCount = list.Count;
				if (list.Count == 0)
				{
					chartResult.Issues.Add("No <c:ser> (series) elements in " + text2 + " chart");
					continue;
				}
				foreach (XElement item in list)
				{
					XElement xElement4 = item.Descendants(ChartNs + "cat").FirstOrDefault();
					if (xElement4 != null)
					{
						XElement xElement5 = xElement4.Descendants(ChartNs + "strRef").FirstOrDefault();
						if (xElement5 != null)
						{
							XElement xElement6 = xElement5.Descendants(ChartNs + "f").FirstOrDefault();
							if (xElement6 != null && !string.IsNullOrEmpty(xElement6.Value) && xElement6.Value != "[0]!categories")
							{
								chartResult.HasCategories = true;
							}
							foreach (XElement item2 in xElement5.Descendants(ChartNs + "strCache").FirstOrDefault()?.Descendants(ChartNs + "pt").ToList() ?? new List<XElement>())
							{
								XElement xElement7 = item2.Element(ChartNs + "v");
								if (xElement7 != null && !string.IsNullOrWhiteSpace(xElement7.Value))
								{
									chartResult.CategoryCount++;
								}
							}
							if (chartResult.CategoryCount > 0)
							{
								chartResult.HasCategories = true;
							}
						}
						XElement xElement8 = xElement4.Descendants(ChartNs + "strLit").FirstOrDefault();
						if (xElement8 != null)
						{
							foreach (XElement item3 in xElement8.Descendants(ChartNs + "pt").ToList())
							{
								XElement xElement9 = item3.Element(ChartNs + "v");
								if (xElement9 != null && !string.IsNullOrWhiteSpace(xElement9.Value))
								{
									chartResult.CategoryCount++;
								}
							}
							if (chartResult.CategoryCount > 0)
							{
								chartResult.HasCategories = true;
							}
						}
						XElement xElement10 = xElement4.Descendants(ChartNs + "numRef").FirstOrDefault();
						if (xElement10 != null)
						{
							XElement xElement11 = xElement10.Descendants(ChartNs + "f").FirstOrDefault();
							if (xElement11 != null && !string.IsNullOrEmpty(xElement11.Value) && xElement11.Value != "[0]!categories")
							{
								chartResult.HasCategories = true;
							}
							List<XElement> list2 = xElement10.Descendants(ChartNs + "numCache").FirstOrDefault()?.Descendants(ChartNs + "pt").ToList() ?? new List<XElement>();
							if (list2.Count > 0)
							{
								chartResult.HasCategories = true;
								chartResult.CategoryCount += list2.Count;
							}
						}
						XElement xElement12 = xElement4.Descendants(ChartNs + "numLit").FirstOrDefault();
						if (xElement12 != null)
						{
							List<XElement> list3 = xElement12.Descendants(ChartNs + "pt").ToList();
							if (list3.Count > 0)
							{
								chartResult.HasCategories = true;
								chartResult.CategoryCount += list3.Count;
							}
						}
					}
					XElement xElement13 = item.Descendants(ChartNs + "val").FirstOrDefault();
					if (xElement13 != null)
					{
						XElement xElement14 = xElement13.Descendants(ChartNs + "numRef").FirstOrDefault();
						if (xElement14 != null)
						{
							XElement xElement15 = xElement14.Descendants(ChartNs + "f").FirstOrDefault();
							if (xElement15 != null && !string.IsNullOrEmpty(xElement15.Value) && xElement15.Value != "0")
							{
								chartResult.HasValues = true;
							}
							List<XElement> list4 = xElement14.Descendants(ChartNs + "numCache").FirstOrDefault()?.Descendants(ChartNs + "pt").ToList() ?? new List<XElement>();
							chartResult.DataPoints += list4.Count;
							foreach (XElement item4 in list4)
							{
								XElement xElement16 = item4.Element(ChartNs + "v");
								if (xElement16 != null && double.TryParse(xElement16.Value, out var result) && result != 0.0)
								{
									chartResult.NonZeroDataPoints++;
								}
							}
						}
						XElement xElement17 = xElement13.Descendants(ChartNs + "numLit").FirstOrDefault();
						if (xElement17 != null)
						{
							List<XElement> list5 = xElement17.Descendants(ChartNs + "pt").ToList();
							chartResult.DataPoints += list5.Count;
							chartResult.HasValues = list5.Count > 0;
							foreach (XElement item5 in list5)
							{
								XElement xElement18 = item5.Element(ChartNs + "v");
								if (xElement18 != null && double.TryParse(xElement18.Value, out var result2) && result2 != 0.0)
								{
									chartResult.NonZeroDataPoints++;
								}
							}
						}
					}
					XElement xElement19 = item.Descendants(ChartNs + "xVal").FirstOrDefault();
					XElement xElement20 = item.Descendants(ChartNs + "yVal").FirstOrDefault();
					if (xElement19 != null || xElement20 != null)
					{
						chartResult.HasValues = true;
					}
				}
				break;
			}
			chartResult.HasMeaningfulData = chartResult.CategoryCount > 0 || chartResult.NonZeroDataPoints > 0;
			if (chartResult.SeriesCount > 0 && (chartResult.HasValues || chartResult.HasCategories))
			{
				chartResult.HasData = true;
				if (!chartResult.HasMeaningfulData)
				{
					chartResult.Issues.Add("WARNING: Chart structure exists but contains no meaningful data (all zeros or empty)");
				}
			}
			else
			{
				if (chartResult.SeriesCount == 0)
				{
					chartResult.Issues.Add("No series found in chart");
				}
				if (!chartResult.HasValues && !chartResult.HasCategories)
				{
					chartResult.Issues.Add("No category or value references found");
				}
			}
		}
		catch (Exception ex)
		{
			chartResult.Issues.Add("Error: " + ex.Message);
		}
		return chartResult;
	}

	public static void PrintReport(VerificationResult results, bool verbose = false)
	{
		Console.WriteLine(new string('=', 60));
		Console.WriteLine("Chart Data Verification Report");
		Console.WriteLine(new string('=', 60));
		Console.WriteLine("File: " + results.File);
		Console.WriteLine($"Total charts: {results.TotalCharts}");
		Console.WriteLine($"Valid charts (with data): {results.ValidCharts.Count}");
		Console.WriteLine($"Charts with NO DATA: {results.ChartsWithNoData.Count}");
		Console.WriteLine($"Empty/broken charts: {results.EmptyCharts.Count}");
		Console.WriteLine(new string('=', 60));
		if (results.Errors.Count > 0)
		{
			Console.WriteLine("\nERRORS:");
			foreach (string error in results.Errors)
			{
				Console.WriteLine("   - " + error);
			}
		}
		if (results.ChartsWithNoData.Count > 0)
		{
			Console.WriteLine("\n⚠\ufe0f  CHARTS WITH NO MEANINGFUL DATA:");
			Console.WriteLine("   (These charts exist but contain no actual data - they will appear empty in Excel)");
			foreach (ChartResult chartsWithNoDatum in results.ChartsWithNoData)
			{
				Console.WriteLine("\n   ❌ " + chartsWithNoDatum.Name);
				Console.WriteLine("      Type: " + (chartsWithNoDatum.ChartType ?? "unknown"));
				Console.WriteLine($"      Series count: {chartsWithNoDatum.SeriesCount}");
				Console.WriteLine($"      Categories: {chartsWithNoDatum.CategoryCount}");
				Console.WriteLine($"      Data points: {chartsWithNoDatum.DataPoints} (non-zero: {chartsWithNoDatum.NonZeroDataPoints})");
				if (chartsWithNoDatum.Issues.Count <= 0)
				{
					continue;
				}
				Console.WriteLine("      Issues:");
				foreach (string issue in chartsWithNoDatum.Issues)
				{
					Console.WriteLine("         - " + issue);
				}
			}
		}
		if (results.EmptyCharts.Count > 0)
		{
			Console.WriteLine("\n❌ EMPTY/BROKEN CHARTS:");
			foreach (ChartResult emptyChart in results.EmptyCharts)
			{
				Console.WriteLine("\n   " + emptyChart.Name);
				Console.WriteLine("      Type: " + (emptyChart.ChartType ?? "unknown"));
				Console.WriteLine($"      Series count: {emptyChart.SeriesCount}");
				Console.WriteLine($"      Has categories: {emptyChart.HasCategories}");
				Console.WriteLine($"      Has values: {emptyChart.HasValues}");
				if (emptyChart.Issues.Count <= 0)
				{
					continue;
				}
				Console.WriteLine("      Issues:");
				foreach (string issue2 in emptyChart.Issues)
				{
					Console.WriteLine("         - " + issue2);
				}
			}
		}
		if (verbose && results.ValidCharts.Count > 0)
		{
			Console.WriteLine("\n✅ VALID CHARTS:");
			foreach (ChartResult validChart in results.ValidCharts)
			{
				Console.WriteLine("\n   " + validChart.Name);
				Console.WriteLine("      Type: " + validChart.ChartType);
				Console.WriteLine($"      Series count: {validChart.SeriesCount}");
				Console.WriteLine($"      Categories: {validChart.CategoryCount}");
				Console.WriteLine($"      Data points: {validChart.DataPoints} (non-zero: {validChart.NonZeroDataPoints})");
			}
		}
		Console.WriteLine("\n" + new string('=', 60));
		if (results.Valid)
		{
			Console.WriteLine("RESULT: ✅ All charts have meaningful data");
		}
		else if (results.ChartsWithNoData.Count > 0)
		{
			Console.WriteLine("RESULT: ❌ FAILED - Some charts have no meaningful data (empty charts)");
		}
		else
		{
			Console.WriteLine("RESULT: ❌ FAILED - Some charts are empty or have issues");
		}
		Console.WriteLine(new string('=', 60) + "\n");
	}
}
