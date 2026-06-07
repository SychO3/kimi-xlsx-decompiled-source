using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace KimiXlsx.Commands;

public static class PivotCommand
{
	private enum CellDataType
	{
		String,
		Number
	}

	private static readonly Dictionary<string, string> PivotStyles = new Dictionary<string, string>
	{
		["monochrome"] = "PivotStyleMedium9",
		["finance"] = "PivotStyleMedium2"
	};

	public static int Run(string[] args)
	{
		if (args.Length < 2)
		{
			PrintUsage();
			return 1;
		}
		string sourceFileName = args[0];
		string text = args[1];
		string text2 = null;
		string text3 = null;
		string rowFields = null;
		string colFields = null;
		string text4 = null;
		string filterFields = null;
		string text5 = "PivotTable1";
		string text6 = "monochrome";
		string text7 = "bar";
		for (int i = 2; i < args.Length; i++)
		{
			switch (args[i])
			{
			case "--source":
				if (i + 1 < args.Length)
				{
					text2 = args[++i];
				}
				break;
			case "--location":
				if (i + 1 < args.Length)
				{
					text3 = args[++i];
				}
				break;
			case "--rows":
				if (i + 1 < args.Length)
				{
					rowFields = args[++i];
				}
				break;
			case "--cols":
				if (i + 1 < args.Length)
				{
					colFields = args[++i];
				}
				break;
			case "--values":
				if (i + 1 < args.Length)
				{
					text4 = args[++i];
				}
				break;
			case "--filters":
				if (i + 1 < args.Length)
				{
					filterFields = args[++i];
				}
				break;
			case "--name":
				if (i + 1 < args.Length)
				{
					text5 = args[++i];
				}
				break;
			case "--style":
				if (i + 1 < args.Length)
				{
					text6 = args[++i].ToLower();
				}
				break;
			case "--chart":
				if (i + 1 < args.Length)
				{
					text7 = args[++i].ToLower();
				}
				break;
			}
		}
		if (string.IsNullOrEmpty(text2))
		{
			Console.WriteLine("Error: --source is required (e.g., --source \"Sheet1!A1:F100\")");
			return 1;
		}
		if (string.IsNullOrEmpty(text3))
		{
			Console.WriteLine("Error: --location is required (e.g., --location \"Summary!A3\")");
			return 1;
		}
		if (string.IsNullOrEmpty(text4))
		{
			Console.WriteLine("Error: --values is required (e.g., --values \"Revenue:sum\")");
			return 1;
		}
		try
		{
			File.Copy(sourceFileName, text, overwrite: true);
			(bool, string) tuple = CreatePivotTable(text, text2, text3, rowFields, colFields, text4, filterFields, text5, text6, text7);
			if (tuple.Item1)
			{
				FixPivotCachePaths(text);
				Console.WriteLine("✅ PivotTable '" + text5 + "' created successfully");
				Console.WriteLine("   Source: " + text2);
				Console.WriteLine("   Location: " + text3);
				Console.WriteLine("   Style: " + text6);
				Console.WriteLine("   Chart: " + text7);
				Console.WriteLine("   Output: " + text);
				return 0;
			}
			Console.WriteLine("❌ Failed to create PivotTable: " + tuple.Item2);
			return 1;
		}
		catch (Exception ex)
		{
			Console.WriteLine("❌ Error: " + ex.Message);
			return 1;
		}
	}

	private static void PrintUsage()
	{
		Console.WriteLine("\nUsage: dotnet run --project KimiXlsx -- pivot <input.xlsx> <output.xlsx> [options]\n\nRequired Options:\n  --source \"Sheet!A1:Z100\"    Source data range\n  --location \"Sheet!A3\"       Where to place the PivotTable\n  --values \"Field:func,...\"   Value fields with aggregation (sum/count/avg/max/min)\n\nOptional:\n  --rows \"Field1,Field2\"      Row fields (comma-separated)\n  --cols \"Field1\"             Column fields (comma-separated)\n  --filters \"Field1\"          Filter fields (comma-separated)\n  --name \"MyPivot\"            PivotTable name (default: PivotTable1)\n  --style \"monochrome|finance\" Style theme (default: monochrome)\n                               monochrome = Black/White/Grey (general analysis)\n                               finance = Blue/White (financial reports)\n  --chart \"bar|line|pie\"      Chart type (default: bar)\n                               bar = Column/Bar chart (category comparison)\n                               line = Line chart (trends over time)\n                               pie = Pie chart (proportion/percentage)\n\nAggregation Functions: sum, count, average, max, min\n\nExamples:\n  dotnet run --project KimiXlsx -- pivot data.xlsx output.xlsx \\\n      --source \"Sales!A1:F100\" --rows \"Product\" --values \"Revenue:sum\" \\\n      --location \"Summary!A3\"\n\n  dotnet run --project KimiXlsx -- pivot data.xlsx output.xlsx \\\n      --source \"Data!A1:H200\" --rows \"Category\" --cols \"Quarter\" \\\n      --values \"Amount:sum,Count:count\" --filters \"Year\" \\\n      --location \"Pivot!B2\" --name \"SalesPivot\" --style \"finance\" --chart \"line\"\n");
	}

	private static (bool Success, string? Error) CreatePivotTable(string filePath, string sourceRange, string location, string? rowFields, string? colFields, string? valueFields, string? filterFields, string pivotName, string styleName, string chartType)
	{
		string builtInStyleName = (PivotStyles.ContainsKey(styleName) ? PivotStyles[styleName] : PivotStyles["monochrome"]);
		Match match = Regex.Match(sourceRange, "^(.+)!([A-Z]+\\d+):([A-Z]+\\d+)$", RegexOptions.IgnoreCase);
		if (!match.Success)
		{
			return (Success: false, Error: "Invalid source range format: " + sourceRange + ". Expected: Sheet!A1:Z100");
		}
		string value = match.Groups[1].Value;
		string text = match.Groups[2].Value + ":" + match.Groups[3].Value;
		Match match2 = Regex.Match(location, "^(.+)!([A-Z]+\\d+)$", RegexOptions.IgnoreCase);
		if (!match2.Success)
		{
			return (Success: false, Error: "Invalid location format: " + location + ". Expected: Sheet!A1");
		}
		string value2 = match2.Groups[1].Value;
		string value3 = match2.Groups[2].Value;
		using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, isEditable: true);
		WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
		WorksheetPart worksheetPart = GetWorksheetPart(workbookPart, value);
		if (worksheetPart == null)
		{
			return (Success: false, Error: "Source sheet '" + value + "' not found");
		}
		var (list, dataRows, dataTypes) = ReadSourceData(worksheetPart, text, workbookPart);
		if (list.Count == 0)
		{
			return (Success: false, Error: "No headers found in source data");
		}
		List<string> list2 = ParseFieldList(rowFields);
		List<string> list3 = ParseFieldList(colFields);
		List<(string, DataConsolidateFunctionValues)> list4 = ParseValueFields(valueFields);
		List<string> list5 = ParseFieldList(filterFields);
		foreach (string item in list2.Concat(list3).Concat(list4.Select<(string, DataConsolidateFunctionValues), string>(((string Field, DataConsolidateFunctionValues Function) v) => v.Field)).Concat(list5))
		{
			if (!list.Contains<string>(item, StringComparer.OrdinalIgnoreCase))
			{
				return (Success: false, Error: "Field '" + item + "' not found in headers: " + string.Join(", ", list));
			}
		}
		WorksheetPart orCreateWorksheetPart = GetOrCreateWorksheetPart(workbookPart, value2);
		try
		{
			CreatePivotTableParts(workbookPart, orCreateWorksheetPart, value, text, value3, pivotName, list, dataRows, dataTypes, list2, list3, list4, list5, builtInStyleName, chartType);
			workbookPart.Workbook.Save();
			return (Success: true, Error: null);
		}
		catch (Exception ex)
		{
			return (Success: false, Error: ex.Message);
		}
	}

	private static WorksheetPart? GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
	{
		Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(delegate(Sheet s)
		{
			StringValue? name = s.Name;
			return (object)name != null && name.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true;
		});
		if (sheet == null || sheet.Id?.Value == null)
		{
			return null;
		}
		return (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
	}

	private static WorksheetPart GetOrCreateWorksheetPart(WorkbookPart workbookPart, string sheetName)
	{
		WorksheetPart worksheetPart = GetWorksheetPart(workbookPart, sheetName);
		if (worksheetPart != null)
		{
			return worksheetPart;
		}
		WorksheetPart worksheetPart2 = workbookPart.AddNewPart<WorksheetPart>();
		SheetViews sheetViews = new SheetViews();
		sheetViews.Append(new SheetView
		{
			ShowGridLines = false,
			WorkbookViewId = 0u
		});
		worksheetPart2.Worksheet = new Worksheet();
		worksheetPart2.Worksheet.Append(sheetViews);
		worksheetPart2.Worksheet.Append(new SheetData());
		Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>() ?? workbookPart.Workbook.AppendChild(new Sheets());
		uint num = 1u;
		if (sheets.Elements<Sheet>().Any())
		{
			num = sheets.Elements<Sheet>().Max((Sheet s) => s.SheetId?.Value ?? 0) + 1;
		}
		Sheet sheet = new Sheet
		{
			Id = workbookPart.GetIdOfPart(worksheetPart2),
			SheetId = num,
			Name = sheetName
		};
		sheets.Append(sheet);
		return worksheetPart2;
	}

	private static (List<string> Headers, List<List<object>> DataRows, List<CellDataType> DataTypes) ReadSourceData(WorksheetPart worksheetPart, string range, WorkbookPart workbookPart)
	{
		List<string> list = new List<string>();
		List<List<object>> list2 = new List<List<object>>();
		List<CellDataType> list3 = new List<CellDataType>();
		SheetData firstChild = worksheetPart.Worksheet.GetFirstChild<SheetData>();
		if (firstChild == null)
		{
			return (Headers: list, DataRows: list2, DataTypes: list3);
		}
		Match match = Regex.Match(range, "([A-Z]+)(\\d+):([A-Z]+)(\\d+)", RegexOptions.IgnoreCase);
		if (!match.Success)
		{
			return (Headers: list, DataRows: list2, DataTypes: list3);
		}
		int num = ColumnToIndex(match.Groups[1].Value);
		int num2 = int.Parse(match.Groups[2].Value);
		int num3 = ColumnToIndex(match.Groups[3].Value);
		int num4 = int.Parse(match.Groups[4].Value);
		List<string> sharedStringsList = LoadSharedStrings(workbookPart);
		bool flag = true;
		foreach (Row item in firstChild.Elements<Row>())
		{
			int num5 = (int)(item.RowIndex?.Value ?? 0);
			if (num5 < num2 || num5 > num4)
			{
				continue;
			}
			List<object> list4 = new List<object>();
			for (int i = num; i <= num3; i++)
			{
				string cellRef = $"{IndexToColumn(i)}{num5}";
				object cellValue = GetCellValue(item.Elements<Cell>().FirstOrDefault((Cell c) => c.CellReference?.Value == cellRef), sharedStringsList);
				if (flag)
				{
					list.Add(cellValue?.ToString() ?? $"Column{i - num + 1}");
				}
				else
				{
					list4.Add(cellValue ?? "");
				}
			}
			if (flag)
			{
				for (int num6 = 0; num6 < list.Count; num6++)
				{
					list3.Add(CellDataType.String);
				}
				flag = false;
				continue;
			}
			list2.Add(list4);
			for (int num7 = 0; num7 < list4.Count && num7 < list3.Count; num7++)
			{
				if (list4[num7] is double || list4[num7] is decimal || list4[num7] is int)
				{
					list3[num7] = CellDataType.Number;
				}
			}
		}
		return (Headers: list, DataRows: list2, DataTypes: list3);
	}

	private static List<string> LoadSharedStrings(WorkbookPart workbookPart)
	{
		List<string> list = new List<string>();
		SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;
		if (sharedStringTablePart == null)
		{
			return list;
		}
		SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;
		if (sharedStringTable == null)
		{
			return list;
		}
		foreach (SharedStringItem item in sharedStringTable.Elements<SharedStringItem>())
		{
			list.Add(item.InnerText ?? "");
		}
		return list;
	}

	private static object? GetCellValue(Cell? cell, List<string> sharedStringsList)
	{
		if (cell == null)
		{
			return null;
		}
		if (cell.DataType?.Value == CellValues.InlineString)
		{
			return cell.InlineString?.InnerText;
		}
		if (cell.CellValue == null)
		{
			return null;
		}
		string text = cell.CellValue.Text;
		CellValues? cellValues = cell.DataType?.Value;
		CellValues sharedString = CellValues.SharedString;
		if (cellValues.HasValue && (!cellValues.HasValue || cellValues.GetValueOrDefault() == sharedString) && int.TryParse(text, out var result))
		{
			if (result >= 0 && result < sharedStringsList.Count)
			{
				return sharedStringsList[result];
			}
			return null;
		}
		if (cell.DataType?.Value == CellValues.String)
		{
			return text;
		}
		if (double.TryParse(text, out var result2))
		{
			return result2;
		}
		return text;
	}

	private static List<string> ParseFieldList(string? fields)
	{
		if (string.IsNullOrEmpty(fields))
		{
			return new List<string>();
		}
		return (from f in fields.Split(',')
			select f.Trim() into f
			where !string.IsNullOrEmpty(f)
			select f).ToList();
	}

	private static List<(string Field, DataConsolidateFunctionValues Function)> ParseValueFields(string? valueFields)
	{
		List<(string, DataConsolidateFunctionValues)> list = new List<(string, DataConsolidateFunctionValues)>();
		if (string.IsNullOrEmpty(valueFields))
		{
			return list;
		}
		string[] array = valueFields.Split(',');
		for (int i = 0; i < array.Length; i++)
		{
			string[] array2 = array[i].Trim().Split(':');
			if (array2.Length >= 2)
			{
				DataConsolidateFunctionValues dataConsolidateFunctionValues;
				switch (array2[1].ToLower())
				{
				case "sum":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.Sum;
					break;
				case "count":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.Count;
					break;
				case "avg":
				case "average":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.Average;
					break;
				case "max":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.Maximum;
					break;
				case "min":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.Minimum;
					break;
				case "product":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.Product;
					break;
				case "countnums":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.CountNumbers;
					break;
				case "stddev":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.StandardDeviation;
					break;
				case "var":
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.Variance;
					break;
				default:
					dataConsolidateFunctionValues = DataConsolidateFunctionValues.Sum;
					break;
				}
				DataConsolidateFunctionValues item = dataConsolidateFunctionValues;
				list.Add((array2[0].Trim(), item));
			}
			else if (array2.Length == 1)
			{
				list.Add((array2[0].Trim(), DataConsolidateFunctionValues.Sum));
			}
		}
		return list;
	}

	private static void CreatePivotTableParts(WorkbookPart workbookPart, WorksheetPart pivotSheetPart, string sourceSheet, string sourceRef, string pivotCell, string pivotName, List<string> headers, List<List<object>> dataRows, List<CellDataType> dataTypes, List<string> rowFields, List<string> colFields, List<(string Field, DataConsolidateFunctionValues Function)> valueFields, List<string> filterFields, string builtInStyleName, string chartType)
	{
		uint num = 1u;
		List<PivotTableCacheDefinitionPart> list = workbookPart.GetPartsOfType<PivotTableCacheDefinitionPart>().ToList();
		if (list.Any())
		{
			foreach (PivotTableCacheDefinitionPart item3 in list)
			{
				Match match = Regex.Match(item3.Uri.ToString(), "pivotCacheDefinition(\\d+)\\.xml");
				if (match.Success && uint.TryParse(match.Groups[1].Value, out var result))
				{
					num = Math.Max(num, result + 1);
				}
			}
		}
		string id = $"rIdPivotCache{num}";
		PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>(id);
		PivotCacheDefinition pivotCacheDefinition = new PivotCacheDefinition
		{
			SaveData = true,
			RefreshOnLoad = true,
			BackgroundQuery = false,
			CreatedVersion = (byte)6,
			RefreshedVersion = (byte)6,
			MinRefreshableVersion = (byte)3,
			RecordCount = (uint)dataRows.Count
		};
		CacheSource cacheSource = new CacheSource
		{
			Type = SourceValues.Worksheet
		};
		cacheSource.Append(new WorksheetSource
		{
			Reference = sourceRef,
			Sheet = sourceSheet
		});
		pivotCacheDefinition.Append(cacheSource);
		Dictionary<string, int> dictionary = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
		for (int i = 0; i < headers.Count; i++)
		{
			dictionary[headers[i]] = i;
		}
		CacheFields cacheFields = new CacheFields
		{
			Count = (uint)headers.Count
		};
		Dictionary<int, List<string>> dictionary2 = new Dictionary<int, List<string>>();
		HashSet<int> hashSet = new HashSet<int>();
		foreach (string rowField in rowFields)
		{
			if (dictionary.TryGetValue(rowField, out var value))
			{
				hashSet.Add(value);
			}
		}
		foreach (string colField in colFields)
		{
			if (dictionary.TryGetValue(colField, out var value2))
			{
				hashSet.Add(value2);
			}
		}
		foreach (string filterField in filterFields)
		{
			if (dictionary.TryGetValue(filterField, out var value3))
			{
				hashSet.Add(value3);
			}
		}
		for (int j = 0; j < headers.Count; j++)
		{
			CacheField cacheField = new CacheField
			{
				Name = headers[j],
				NumberFormatId = 0u
			};
			bool flag = hashSet.Contains(j);
			if (dataTypes[j] == CellDataType.Number && !flag)
			{
				double num2 = double.MaxValue;
				double num3 = double.MinValue;
				foreach (List<object> dataRow in dataRows)
				{
					if (j < dataRow.Count && dataRow[j] is double val)
					{
						num2 = Math.Min(num2, val);
						num3 = Math.Max(num3, val);
					}
				}
				SharedItems sharedItems = new SharedItems
				{
					ContainsSemiMixedTypes = false,
					ContainsString = false,
					ContainsNumber = true,
					MinValue = ((num2 == double.MaxValue) ? 0.0 : num2),
					MaxValue = ((num3 == double.MinValue) ? 0.0 : num3)
				};
				cacheField.Append(sharedItems);
			}
			else if (dataTypes[j] == CellDataType.Number && flag)
			{
				HashSet<string> hashSet2 = new HashSet<string>();
				foreach (List<object> dataRow2 in dataRows)
				{
					if (j < dataRow2.Count)
					{
						object obj = dataRow2[j];
						string text = "";
						text = ((!(obj is double num4)) ? ((!(obj is int num5)) ? (obj?.ToString() ?? "") : num5.ToString()) : ((int)num4).ToString());
						hashSet2.Add(text);
					}
				}
				List<string> list2 = (dictionary2[j] = hashSet2.OrderBy((string s) => s).ToList());
				SharedItems sharedItems2 = new SharedItems
				{
					Count = (uint)list2.Count
				};
				foreach (string item4 in list2)
				{
					sharedItems2.Append(new StringItem
					{
						Val = item4
					});
				}
				cacheField.Append(sharedItems2);
			}
			else
			{
				HashSet<string> hashSet3 = new HashSet<string>();
				foreach (List<object> dataRow3 in dataRows)
				{
					if (j < dataRow3.Count)
					{
						hashSet3.Add(dataRow3[j]?.ToString() ?? "");
					}
				}
				List<string> list4 = (dictionary2[j] = hashSet3.OrderBy((string s) => s).ToList());
				SharedItems sharedItems3 = new SharedItems
				{
					Count = (uint)list4.Count
				};
				foreach (string item5 in list4)
				{
					sharedItems3.Append(new StringItem
					{
						Val = item5
					});
				}
				cacheField.Append(sharedItems3);
			}
			cacheFields.Append(cacheField);
		}
		pivotCacheDefinition.Append(cacheFields);
		pivotTableCacheDefinitionPart.PivotCacheDefinition = pivotCacheDefinition;
		string id2 = $"rIdPivotCacheRecords{num}";
		PivotTableCacheRecordsPart pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>(id2);
		PivotCacheRecords pivotCacheRecords = new PivotCacheRecords
		{
			Count = (uint)dataRows.Count
		};
		foreach (List<object> dataRow4 in dataRows)
		{
			PivotCacheRecord pivotCacheRecord = new PivotCacheRecord();
			for (int num6 = 0; num6 < headers.Count && num6 < dataRow4.Count; num6++)
			{
				bool flag2 = dataTypes[num6] == CellDataType.Number && hashSet.Contains(num6);
				if (dataTypes[num6] == CellDataType.Number && !flag2)
				{
					double result2 = 0.0;
					if (dataRow4[num6] is double num7)
					{
						result2 = num7;
					}
					else if (dataRow4[num6] is int num8)
					{
						result2 = num8;
					}
					else
					{
						double.TryParse(dataRow4[num6]?.ToString(), out result2);
					}
					pivotCacheRecord.Append(new NumberItem
					{
						Val = result2
					});
				}
				else
				{
					string text2 = "";
					text2 = ((!(dataRow4[num6] is double num9)) ? ((!(dataRow4[num6] is int num10)) ? (dataRow4[num6]?.ToString() ?? "") : num10.ToString()) : ((int)num9).ToString());
					int num11 = (dictionary2.ContainsKey(num6) ? dictionary2[num6].IndexOf(text2) : 0);
					if (num11 < 0)
					{
						num11 = 0;
					}
					pivotCacheRecord.Append(new FieldItem
					{
						Val = (uint)num11
					});
				}
			}
			pivotCacheRecords.Append(pivotCacheRecord);
		}
		pivotTableCacheRecordsPart.PivotCacheRecords = pivotCacheRecords;
		PivotCaches pivotCaches = workbookPart.Workbook.GetFirstChild<PivotCaches>();
		if (pivotCaches == null)
		{
			pivotCaches = new PivotCaches();
			CalculationProperties firstChild = workbookPart.Workbook.GetFirstChild<CalculationProperties>();
			Sheets firstChild2 = workbookPart.Workbook.GetFirstChild<Sheets>();
			if (firstChild != null)
			{
				workbookPart.Workbook.InsertAfter(pivotCaches, firstChild);
			}
			else if (firstChild2 != null)
			{
				workbookPart.Workbook.InsertAfter(pivotCaches, firstChild2);
			}
			else
			{
				workbookPart.Workbook.Append(pivotCaches);
			}
		}
		uint num12 = 1u;
		if (pivotCaches.Elements<PivotCache>().Any())
		{
			num12 = pivotCaches.Elements<PivotCache>().Max((PivotCache p) => p.CacheId?.Value ?? 0) + 1;
		}
		pivotCaches.Append(new PivotCache
		{
			CacheId = num12,
			Id = workbookPart.GetIdOfPart(pivotTableCacheDefinitionPart)
		});
		PivotTablePart pivotTablePart = pivotSheetPart.AddNewPart<PivotTablePart>();
		pivotTablePart.AddPart(pivotTableCacheDefinitionPart);
		(int Col, int Row) tuple = ParseCellReference(pivotCell);
		int item = tuple.Col;
		int item2 = tuple.Row;
		int count = rowFields.Count;
		int count2 = colFields.Count;
		int count3 = valueFields.Count;
		int num13 = Math.Min(50, dataRows.Count + 5);
		int num14 = count + Math.Max(1, count2) * count3 + 2;
		string text3 = $"{IndexToColumn(item + num14 - 1)}{item2 + num13}";
		PivotTableDefinition pivotTableDefinition = new PivotTableDefinition
		{
			Name = pivotName,
			CacheId = num12,
			DataCaption = "Values",
			UpdatedVersion = (byte)6,
			MinRefreshableVersion = (byte)3,
			UseAutoFormatting = true,
			ItemPrintTitles = true,
			CreatedVersion = (byte)6,
			Indent = 0u,
			Outline = true,
			OutlineData = true,
			MultipleFieldFilters = false,
			ApplyNumberFormats = false,
			ApplyBorderFormats = false,
			ApplyFontFormats = false,
			ApplyPatternFormats = false,
			ApplyAlignmentFormats = false,
			ApplyWidthHeightFormats = true
		};
		pivotTableDefinition.Append(new Location
		{
			Reference = pivotCell + ":" + text3,
			FirstHeaderRow = 1u,
			FirstDataRow = ((count2 <= 0) ? 1u : 2u),
			FirstDataColumn = (uint)count
		});
		PivotFields pivotFields = new PivotFields
		{
			Count = (uint)headers.Count
		};
		for (int num15 = 0; num15 < headers.Count; num15++)
		{
			string fieldName = headers[num15];
			PivotField pivotField = new PivotField
			{
				ShowAll = false
			};
			bool num16 = rowFields.Any((string f) => f.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
			bool flag3 = colFields.Any((string f) => f.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
			bool flag4 = valueFields.Any<(string, DataConsolidateFunctionValues)>(((string Field, DataConsolidateFunctionValues Function) f) => f.Field.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
			bool flag5 = filterFields.Any((string f) => f.Equals(fieldName, StringComparison.OrdinalIgnoreCase));
			if (num16)
			{
				pivotField.Axis = PivotTableAxisValues.AxisRow;
				if (dictionary2.ContainsKey(num15))
				{
					Items items = new Items
					{
						Count = (uint)(dictionary2[num15].Count + 1)
					};
					for (int num17 = 0; num17 < dictionary2[num15].Count; num17++)
					{
						items.Append(new Item
						{
							Index = (uint)num17
						});
					}
					items.Append(new Item
					{
						ItemType = ItemValues.Default
					});
					pivotField.Append(items);
				}
			}
			else if (flag3)
			{
				pivotField.Axis = PivotTableAxisValues.AxisColumn;
				if (dictionary2.ContainsKey(num15))
				{
					Items items2 = new Items
					{
						Count = (uint)(dictionary2[num15].Count + 1)
					};
					for (int num18 = 0; num18 < dictionary2[num15].Count; num18++)
					{
						items2.Append(new Item
						{
							Index = (uint)num18
						});
					}
					items2.Append(new Item
					{
						ItemType = ItemValues.Default
					});
					pivotField.Append(items2);
				}
			}
			else if (flag5)
			{
				pivotField.Axis = PivotTableAxisValues.AxisPage;
				if (dictionary2.ContainsKey(num15))
				{
					Items items3 = new Items
					{
						Count = (uint)(dictionary2[num15].Count + 1)
					};
					for (int num19 = 0; num19 < dictionary2[num15].Count; num19++)
					{
						items3.Append(new Item
						{
							Index = (uint)num19
						});
					}
					items3.Append(new Item
					{
						ItemType = ItemValues.Default
					});
					pivotField.Append(items3);
				}
			}
			else if (flag4)
			{
				pivotField.DataField = true;
			}
			pivotFields.Append(pivotField);
		}
		pivotTableDefinition.Append(pivotFields);
		if (rowFields.Count > 0)
		{
			RowFields rowFields2 = new RowFields
			{
				Count = (uint)rowFields.Count
			};
			foreach (string rowField2 in rowFields)
			{
				if (dictionary.TryGetValue(rowField2, out var value4))
				{
					rowFields2.Append(new DocumentFormat.OpenXml.Spreadsheet.Field
					{
						Index = value4
					});
				}
			}
			pivotTableDefinition.Append(rowFields2);
			int num20 = 1;
			if (rowFields.Count > 0 && dictionary.TryGetValue(rowFields[0], out var value5) && dictionary2.ContainsKey(value5))
			{
				num20 = dictionary2[value5].Count + 1;
			}
			RowItems rowItems = new RowItems
			{
				Count = (uint)num20
			};
			for (int num21 = 0; num21 < num20 - 1; num21++)
			{
				RowItem rowItem = new RowItem();
				rowItem.Append(new MemberPropertyIndex
				{
					Val = num21
				});
				rowItems.Append(rowItem);
			}
			RowItem rowItem2 = new RowItem
			{
				ItemType = ItemValues.Grand
			};
			rowItem2.Append(new MemberPropertyIndex());
			rowItems.Append(rowItem2);
			pivotTableDefinition.Append(rowItems);
		}
		if (colFields.Count > 0)
		{
			ColumnFields columnFields = new ColumnFields
			{
				Count = (uint)colFields.Count
			};
			foreach (string colField2 in colFields)
			{
				if (dictionary.TryGetValue(colField2, out var value6))
				{
					columnFields.Append(new DocumentFormat.OpenXml.Spreadsheet.Field
					{
						Index = value6
					});
				}
			}
			pivotTableDefinition.Append(columnFields);
			int num22 = 1;
			if (colFields.Count > 0 && dictionary.TryGetValue(colFields[0], out var value7) && dictionary2.ContainsKey(value7))
			{
				num22 = dictionary2[value7].Count + 1;
			}
			ColumnItems columnItems = new ColumnItems
			{
				Count = (uint)num22
			};
			for (int num23 = 0; num23 < num22 - 1; num23++)
			{
				RowItem rowItem3 = new RowItem();
				rowItem3.Append(new MemberPropertyIndex
				{
					Val = num23
				});
				columnItems.Append(rowItem3);
			}
			RowItem rowItem4 = new RowItem
			{
				ItemType = ItemValues.Grand
			};
			rowItem4.Append(new MemberPropertyIndex());
			columnItems.Append(rowItem4);
			pivotTableDefinition.Append(columnItems);
		}
		else if (valueFields.Count > 1)
		{
			ColumnFields columnFields2 = new ColumnFields
			{
				Count = 1u
			};
			columnFields2.Append(new DocumentFormat.OpenXml.Spreadsheet.Field
			{
				Index = -2
			});
			pivotTableDefinition.Append(columnFields2);
		}
		if (filterFields.Count > 0)
		{
			PageFields pageFields = new PageFields
			{
				Count = (uint)filterFields.Count
			};
			foreach (string filterField2 in filterFields)
			{
				if (dictionary.TryGetValue(filterField2, out var value8))
				{
					pageFields.Append(new PageField
					{
						Field = value8,
						Hierarchy = -1
					});
				}
			}
			pivotTableDefinition.Append(pageFields);
		}
		if (valueFields.Count > 0)
		{
			DataFields dataFields = new DataFields
			{
				Count = (uint)valueFields.Count
			};
			foreach (var valueField in valueFields)
			{
				if (dictionary.TryGetValue(valueField.Field, out var value9))
				{
					string functionName = GetFunctionName(valueField.Function);
					dataFields.Append(new DataField
					{
						Name = functionName + " of " + valueField.Field,
						Field = (uint)value9,
						Subtotal = valueField.Function,
						BaseField = 0,
						BaseItem = 0u
					});
				}
			}
			pivotTableDefinition.Append(dataFields);
		}
		pivotTableDefinition.Append(new PivotTableStyle
		{
			Name = builtInStyleName,
			ShowRowHeaders = true,
			ShowColumnHeaders = true,
			ShowRowStripes = true,
			ShowColumnStripes = false,
			ShowLastColumn = true
		});
		pivotTablePart.PivotTableDefinition = pivotTableDefinition;
		CreateChart(pivotSheetPart, pivotCell, rowFields, valueFields, dictionary2, dictionary, dataRows, dataTypes, chartType);
	}

	private static void CreateChart(WorksheetPart pivotSheetPart, string pivotCell, List<string> rowFields, List<(string Field, DataConsolidateFunctionValues Function)> valueFields, Dictionary<int, List<string>> uniqueValues, Dictionary<string, int> fieldIndexMap, List<List<object>> dataRows, List<CellDataType> dataTypes, string chartType)
	{
		DrawingsPart drawingsPart;
		if (pivotSheetPart.DrawingsPart == null)
		{
			drawingsPart = pivotSheetPart.AddNewPart<DrawingsPart>();
			drawingsPart.WorksheetDrawing = new WorksheetDrawing();
			Drawing drawing = new Drawing
			{
				Id = pivotSheetPart.GetIdOfPart(drawingsPart)
			};
			pivotSheetPart.Worksheet.Append(drawing);
		}
		else
		{
			drawingsPart = pivotSheetPart.DrawingsPart;
		}
		ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
		(int Col, int Row) tuple = ParseCellReference(pivotCell);
		int item = tuple.Col;
		int item2 = tuple.Row;
		int num = item + 5;
		int num2 = item2;
		int num3 = num + 8;
		int num4 = num2 + 15;
		GenerateChartContent(chartPart, rowFields, valueFields, uniqueValues, fieldIndexMap, dataRows, dataTypes, chartType);
		WorksheetDrawing worksheetDrawing = drawingsPart.WorksheetDrawing;
		TwoCellAnchor twoCellAnchor = new TwoCellAnchor
		{
			EditAs = EditAsValues.OneCell
		};
		twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId(num.ToString()), new ColumnOffset("0"), new RowId((num2 - 1).ToString()), new RowOffset("0")));
		twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId(num3.ToString()), new ColumnOffset("0"), new RowId((num4 - 1).ToString()), new RowOffset("0")));
		DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame();
		DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties();
		nonVisualGraphicFrameProperties.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties
		{
			Id = 2u,
			Name = "PivotChart 1"
		});
		nonVisualGraphicFrameProperties.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties());
		graphicFrame.Append(nonVisualGraphicFrameProperties);
		graphicFrame.Append(new Transform(new Offset
		{
			X = 0L,
			Y = 0L
		}, new Extents
		{
			Cx = 0L,
			Cy = 0L
		}));
		Graphic graphic = new Graphic();
		GraphicData graphicData = new GraphicData
		{
			Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
		};
		ChartReference chartReference = new ChartReference
		{
			Id = drawingsPart.GetIdOfPart(chartPart)
		};
		graphicData.Append(chartReference);
		graphic.Append(graphicData);
		graphicFrame.Append(graphic);
		twoCellAnchor.Append(graphicFrame);
		twoCellAnchor.Append(new ClientData());
		worksheetDrawing.Append(twoCellAnchor);
	}

	private static void GenerateChartContent(ChartPart chartPart, List<string> rowFields, List<(string Field, DataConsolidateFunctionValues Function)> valueFields, Dictionary<int, List<string>> uniqueValues, Dictionary<string, int> fieldIndexMap, List<List<object>> dataRows, List<CellDataType> dataTypes, string chartType)
	{
		Dictionary<string, List<double>> dictionary = new Dictionary<string, List<double>>();
		if (rowFields.Count > 0 && fieldIndexMap.TryGetValue(rowFields[0], out var value))
		{
			foreach (List<object> dataRow in dataRows)
			{
				if (value >= dataRow.Count)
				{
					continue;
				}
				string key = dataRow[value]?.ToString() ?? "";
				if (!dictionary.ContainsKey(key))
				{
					dictionary[key] = new List<double>(new double[valueFields.Count]);
				}
				for (int i = 0; i < valueFields.Count; i++)
				{
					(string, DataConsolidateFunctionValues) tuple = valueFields[i];
					if (!fieldIndexMap.TryGetValue(tuple.Item1, out var value2) || value2 >= dataRow.Count)
					{
						continue;
					}
					double result = 0.0;
					if (dataRow[value2] is double num)
					{
						result = num;
					}
					else if (dataRow[value2] is int num2)
					{
						result = num2;
					}
					else
					{
						double.TryParse(dataRow[value2]?.ToString(), out result);
					}
					if (tuple.Item2 == DataConsolidateFunctionValues.Sum || tuple.Item2 == DataConsolidateFunctionValues.Average)
					{
						dictionary[key][i] += result;
					}
					else if (tuple.Item2 == DataConsolidateFunctionValues.Count)
					{
						dictionary[key][i] += 1.0;
					}
					else if (tuple.Item2 == DataConsolidateFunctionValues.Maximum)
					{
						dictionary[key][i] = Math.Max(dictionary[key][i], result);
					}
					else if (tuple.Item2 == DataConsolidateFunctionValues.Minimum)
					{
						if (dictionary[key][i] == 0.0)
						{
							dictionary[key][i] = result;
						}
						else
						{
							dictionary[key][i] = Math.Min(dictionary[key][i], result);
						}
					}
				}
			}
		}
		List<string> categories = dictionary.Keys.OrderBy((string k) => k).ToList();
		ChartSpace chartSpace = new ChartSpace();
		chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
		chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
		DocumentFormat.OpenXml.Drawing.Charts.Chart chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
		chart.Append(new AutoTitleDeleted
		{
			Val = true
		});
		PlotArea plotArea = new PlotArea();
		plotArea.Append(new Layout());
		if (chartType == "pie")
		{
			GeneratePieChart(plotArea, categories, dictionary, valueFields);
		}
		else if (chartType == "line")
		{
			GenerateLineChart(plotArea, categories, dictionary, valueFields);
		}
		else
		{
			GenerateBarChart(plotArea, categories, dictionary, valueFields);
		}
		chart.Append(plotArea);
		Legend legend = new Legend();
		legend.Append(new LegendPosition
		{
			Val = LegendPositionValues.Bottom
		});
		legend.Append(new Overlay
		{
			Val = false
		});
		chart.Append(legend);
		chart.Append(new PlotVisibleOnly
		{
			Val = true
		});
		chart.Append(new DisplayBlanksAs
		{
			Val = DisplayBlanksAsValues.Gap
		});
		chartSpace.Append(chart);
		PrintSettings printSettings = new PrintSettings();
		printSettings.Append(new DocumentFormat.OpenXml.Drawing.Charts.HeaderFooter());
		printSettings.Append(new DocumentFormat.OpenXml.Drawing.Charts.PageMargins
		{
			Left = 0.7,
			Right = 0.7,
			Top = 0.75,
			Bottom = 0.75,
			Header = 0.3,
			Footer = 0.3
		});
		printSettings.Append(new DocumentFormat.OpenXml.Drawing.Charts.PageSetup());
		chartSpace.Append(printSettings);
		chartPart.ChartSpace = chartSpace;
	}

	private static void GenerateBarChart(PlotArea plotArea, List<string> categories, Dictionary<string, List<double>> chartData, List<(string Field, DataConsolidateFunctionValues Function)> valueFields)
	{
		BarChart barChart = new BarChart();
		barChart.Append(new BarDirection
		{
			Val = BarDirectionValues.Column
		});
		barChart.Append(new BarGrouping
		{
			Val = BarGroupingValues.Clustered
		});
		barChart.Append(new VaryColors
		{
			Val = false
		});
		uint num = 0u;
		foreach (var valueField in valueFields)
		{
			BarChartSeries barChartSeries = new BarChartSeries();
			barChartSeries.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index
			{
				Val = num
			});
			barChartSeries.Append(new Order
			{
				Val = num
			});
			SeriesText seriesText = new SeriesText();
			seriesText.Append(new NumericValue(GetFunctionName(valueField.Function) + " of " + valueField.Field));
			barChartSeries.Append(seriesText);
			AppendCategoryAndValues(barChartSeries, categories, chartData, (int)num);
			barChart.Append(barChartSeries);
			num++;
		}
		AppendDataLabels(barChart);
		barChart.Append(new GapWidth
		{
			Val = (ushort)150
		});
		barChart.Append(new AxisId
		{
			Val = 1u
		});
		barChart.Append(new AxisId
		{
			Val = 2u
		});
		plotArea.Append(barChart);
		AppendCategoryAxis(plotArea);
		AppendValueAxis(plotArea);
	}

	private static void GenerateLineChart(PlotArea plotArea, List<string> categories, Dictionary<string, List<double>> chartData, List<(string Field, DataConsolidateFunctionValues Function)> valueFields)
	{
		LineChart lineChart = new LineChart();
		lineChart.Append(new Grouping
		{
			Val = GroupingValues.Standard
		});
		lineChart.Append(new VaryColors
		{
			Val = false
		});
		uint num = 0u;
		foreach (var valueField in valueFields)
		{
			LineChartSeries lineChartSeries = new LineChartSeries();
			lineChartSeries.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index
			{
				Val = num
			});
			lineChartSeries.Append(new Order
			{
				Val = num
			});
			SeriesText seriesText = new SeriesText();
			seriesText.Append(new NumericValue(GetFunctionName(valueField.Function) + " of " + valueField.Field));
			lineChartSeries.Append(seriesText);
			lineChartSeries.Append(new Marker(new Symbol
			{
				Val = MarkerStyleValues.Circle
			}));
			AppendCategoryAndValues(lineChartSeries, categories, chartData, (int)num);
			lineChart.Append(lineChartSeries);
			num++;
		}
		AppendDataLabels(lineChart);
		lineChart.Append(new ShowMarker
		{
			Val = true
		});
		lineChart.Append(new Smooth
		{
			Val = false
		});
		lineChart.Append(new AxisId
		{
			Val = 1u
		});
		lineChart.Append(new AxisId
		{
			Val = 2u
		});
		plotArea.Append(lineChart);
		AppendCategoryAxis(plotArea);
		AppendValueAxis(plotArea);
	}

	private static void GeneratePieChart(PlotArea plotArea, List<string> categories, Dictionary<string, List<double>> chartData, List<(string Field, DataConsolidateFunctionValues Function)> valueFields)
	{
		PieChart pieChart = new PieChart();
		pieChart.Append(new VaryColors
		{
			Val = true
		});
		(string, DataConsolidateFunctionValues) tuple = valueFields[0];
		PieChartSeries pieChartSeries = new PieChartSeries();
		pieChartSeries.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index
		{
			Val = 0u
		});
		pieChartSeries.Append(new Order
		{
			Val = 0u
		});
		SeriesText seriesText = new SeriesText();
		seriesText.Append(new NumericValue(GetFunctionName(tuple.Item2) + " of " + tuple.Item1));
		pieChartSeries.Append(seriesText);
		AppendCategoryAndValues(pieChartSeries, categories, chartData, 0);
		pieChart.Append(pieChartSeries);
		DataLabels dataLabels = new DataLabels();
		dataLabels.Append(new ShowLegendKey
		{
			Val = false
		});
		dataLabels.Append(new ShowValue
		{
			Val = false
		});
		dataLabels.Append(new ShowCategoryName
		{
			Val = true
		});
		dataLabels.Append(new ShowSeriesName
		{
			Val = false
		});
		dataLabels.Append(new ShowPercent
		{
			Val = true
		});
		dataLabels.Append(new ShowBubbleSize
		{
			Val = false
		});
		pieChart.Append(dataLabels);
		pieChart.Append(new FirstSliceAngle
		{
			Val = (ushort)0
		});
		plotArea.Append(pieChart);
	}

	private static void AppendCategoryAndValues(OpenXmlCompositeElement series, List<string> categories, Dictionary<string, List<double>> chartData, int seriesIndex)
	{
		CategoryAxisData categoryAxisData = new CategoryAxisData();
		StringLiteral stringLiteral = new StringLiteral();
		stringLiteral.Append(new PointCount
		{
			Val = (uint)categories.Count
		});
		for (int i = 0; i < categories.Count; i++)
		{
			stringLiteral.Append(new StringPoint(new NumericValue(categories[i]))
			{
				Index = (uint)i
			});
		}
		categoryAxisData.Append(stringLiteral);
		series.Append(categoryAxisData);
		DocumentFormat.OpenXml.Drawing.Charts.Values values = new DocumentFormat.OpenXml.Drawing.Charts.Values();
		NumberLiteral numberLiteral = new NumberLiteral();
		numberLiteral.Append(new FormatCode("General"));
		numberLiteral.Append(new PointCount
		{
			Val = (uint)categories.Count
		});
		for (int j = 0; j < categories.Count; j++)
		{
			double num = chartData[categories[j]][seriesIndex];
			numberLiteral.Append(new NumericPoint(new NumericValue(num.ToString()))
			{
				Index = (uint)j
			});
		}
		values.Append(numberLiteral);
		series.Append(values);
	}

	private static void AppendDataLabels(OpenXmlCompositeElement chartElement)
	{
		DataLabels dataLabels = new DataLabels();
		dataLabels.Append(new ShowLegendKey
		{
			Val = false
		});
		dataLabels.Append(new ShowValue
		{
			Val = false
		});
		dataLabels.Append(new ShowCategoryName
		{
			Val = false
		});
		dataLabels.Append(new ShowSeriesName
		{
			Val = false
		});
		dataLabels.Append(new ShowPercent
		{
			Val = false
		});
		dataLabels.Append(new ShowBubbleSize
		{
			Val = false
		});
		chartElement.Append(dataLabels);
	}

	private static void AppendCategoryAxis(PlotArea plotArea)
	{
		CategoryAxis categoryAxis = new CategoryAxis();
		categoryAxis.Append(new AxisId
		{
			Val = 1u
		});
		categoryAxis.Append(new Scaling(new Orientation
		{
			Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax
		}));
		categoryAxis.Append(new Delete
		{
			Val = false
		});
		categoryAxis.Append(new AxisPosition
		{
			Val = AxisPositionValues.Bottom
		});
		categoryAxis.Append(new MajorTickMark
		{
			Val = TickMarkValues.Outside
		});
		categoryAxis.Append(new MinorTickMark
		{
			Val = TickMarkValues.None
		});
		categoryAxis.Append(new TickLabelPosition
		{
			Val = TickLabelPositionValues.NextTo
		});
		categoryAxis.Append(new CrossingAxis
		{
			Val = 2u
		});
		categoryAxis.Append(new Crosses
		{
			Val = CrossesValues.AutoZero
		});
		categoryAxis.Append(new AutoLabeled
		{
			Val = true
		});
		categoryAxis.Append(new LabelAlignment
		{
			Val = LabelAlignmentValues.Center
		});
		categoryAxis.Append(new LabelOffset
		{
			Val = (ushort)100
		});
		plotArea.Append(categoryAxis);
	}

	private static void AppendValueAxis(PlotArea plotArea)
	{
		ValueAxis valueAxis = new ValueAxis();
		valueAxis.Append(new AxisId
		{
			Val = 2u
		});
		valueAxis.Append(new Scaling(new Orientation
		{
			Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax
		}));
		valueAxis.Append(new Delete
		{
			Val = false
		});
		valueAxis.Append(new AxisPosition
		{
			Val = AxisPositionValues.Left
		});
		valueAxis.Append(new MajorGridlines());
		valueAxis.Append(new MajorTickMark
		{
			Val = TickMarkValues.Outside
		});
		valueAxis.Append(new MinorTickMark
		{
			Val = TickMarkValues.None
		});
		valueAxis.Append(new TickLabelPosition
		{
			Val = TickLabelPositionValues.NextTo
		});
		valueAxis.Append(new CrossingAxis
		{
			Val = 1u
		});
		valueAxis.Append(new Crosses
		{
			Val = CrossesValues.AutoZero
		});
		valueAxis.Append(new CrossBetween
		{
			Val = CrossBetweenValues.Between
		});
		plotArea.Append(valueAxis);
	}

	private static (int Col, int Row) ParseCellReference(string cellRef)
	{
		Match match = Regex.Match(cellRef, "([A-Z]+)(\\d+)", RegexOptions.IgnoreCase);
		if (!match.Success)
		{
			return (Col: 0, Row: 1);
		}
		return (Col: ColumnToIndex(match.Groups[1].Value), Row: int.Parse(match.Groups[2].Value));
	}

	private static string GetFunctionName(DataConsolidateFunctionValues func)
	{
		if (func == DataConsolidateFunctionValues.Sum)
		{
			return "Sum";
		}
		if (func == DataConsolidateFunctionValues.Count)
		{
			return "Count";
		}
		if (func == DataConsolidateFunctionValues.Average)
		{
			return "Average";
		}
		if (func == DataConsolidateFunctionValues.Maximum)
		{
			return "Max";
		}
		if (func == DataConsolidateFunctionValues.Minimum)
		{
			return "Min";
		}
		if (func == DataConsolidateFunctionValues.Product)
		{
			return "Product";
		}
		if (func == DataConsolidateFunctionValues.CountNumbers)
		{
			return "CountNums";
		}
		if (func == DataConsolidateFunctionValues.StandardDeviation)
		{
			return "StdDev";
		}
		if (func == DataConsolidateFunctionValues.Variance)
		{
			return "Var";
		}
		return "Sum";
	}

	private static int ColumnToIndex(string col)
	{
		int num = 0;
		string text = col.ToUpper();
		foreach (char c in text)
		{
			num = num * 26 + (c - 65 + 1);
		}
		return num - 1;
	}

	private static string IndexToColumn(int idx)
	{
		string text = "";
		for (idx++; idx > 0; idx /= 26)
		{
			idx--;
			text = (char)(65 + idx % 26) + text;
		}
		return text;
	}

	private static void FixPivotCachePaths(string filePath)
	{
		string text = filePath + ".tmp";
		try
		{
			using (ZipArchive zipArchive = ZipFile.Open(filePath, ZipArchiveMode.Read))
			{
				using ZipArchive zipArchive2 = ZipFile.Open(text, ZipArchiveMode.Create);
				List<ZipArchiveEntry> list = zipArchive.Entries.ToList();
				Dictionary<string, string> dictionary = new Dictionary<string, string>();
				foreach (ZipArchiveEntry item in list)
				{
					if (item.FullName.StartsWith("pivotCache/", StringComparison.OrdinalIgnoreCase))
					{
						string value = "xl/" + item.FullName;
						dictionary[item.FullName] = value;
					}
				}
				foreach (ZipArchiveEntry item2 in list)
				{
					string entryName = item2.FullName;
					if (dictionary.TryGetValue(item2.FullName, out var value2))
					{
						entryName = value2;
					}
					ZipArchiveEntry zipArchiveEntry = zipArchive2.CreateEntry(entryName);
					using Stream stream = item2.Open();
					using Stream stream2 = zipArchiveEntry.Open();
					if (item2.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
					{
						FixRelationshipFile(stream, stream2, item2.FullName, dictionary);
					}
					else if (item2.FullName.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
					{
						FixContentTypesFile(stream, stream2, dictionary);
					}
					else
					{
						stream.CopyTo(stream2);
					}
				}
			}
			File.Delete(filePath);
			File.Move(text, filePath);
		}
		catch
		{
			if (File.Exists(text))
			{
				File.Delete(text);
			}
			throw;
		}
	}

	private static void FixRelationshipFile(Stream input, Stream output, string relsPath, Dictionary<string, string> pivotCacheMoves)
	{
		XNamespace xNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
		XDocument xDocument = XDocument.Load(input);
		string text = relsPath;
		if (relsPath.StartsWith("pivotCache/", StringComparison.OrdinalIgnoreCase))
		{
			text = "xl/" + relsPath;
		}
		string basePath = "";
		if (text.Contains("/_rels/"))
		{
			basePath = text.Substring(0, text.IndexOf("/_rels/") + 1);
		}
		foreach (XElement item in xDocument.Descendants(xNamespace + "Relationship"))
		{
			string text2 = item.Attribute("Target")?.Value;
			if (text2 != null && text2.StartsWith("/"))
			{
				string text3 = text2.TrimStart('/');
				if (pivotCacheMoves.TryGetValue(text3, out string value))
				{
					text3 = value;
				}
				string value2 = MakeRelativePath(basePath, text3);
				item.SetAttributeValue("Target", value2);
			}
		}
		xDocument.Save(output);
	}

	private static void FixContentTypesFile(Stream input, Stream output, Dictionary<string, string> pivotCacheMoves)
	{
		XNamespace xNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
		XDocument xDocument = XDocument.Load(input);
		foreach (XElement item in xDocument.Descendants(xNamespace + "Override"))
		{
			string text = item.Attribute("PartName")?.Value;
			if (text != null)
			{
				string key = text.TrimStart('/');
				if (pivotCacheMoves.TryGetValue(key, out string value))
				{
					item.SetAttributeValue("PartName", "/" + value);
				}
			}
		}
		xDocument.Save(output);
	}

	private static string MakeRelativePath(string basePath, string targetPath)
	{
		if (string.IsNullOrEmpty(basePath))
		{
			return targetPath;
		}
		List<string> list = (from s in basePath.TrimEnd('/').Split('/')
			where !string.IsNullOrEmpty(s)
			select s).ToList();
		List<string> list2 = (from s in targetPath.Split('/')
			where !string.IsNullOrEmpty(s)
			select s).ToList();
		int num;
		for (num = 0; num < list.Count && num < list2.Count && list[num].Equals(list2[num], StringComparison.OrdinalIgnoreCase); num++)
		{
		}
		List<string> list3 = new List<string>();
		for (int num2 = num; num2 < list.Count; num2++)
		{
			list3.Add("..");
		}
		for (int num3 = num; num3 < list2.Count; num3++)
		{
			list3.Add(list2[num3]);
		}
		return string.Join("/", list3);
	}
}
