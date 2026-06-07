using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using ClosedXML.Excel;

namespace KimiXlsx.Commands;

public static class InspectCommand
{
	private static readonly string[] TotalKeywords = new string[7] { "total", "sum", "grand total", "合计", "总计", "小计", "总结" };

	private static readonly string[] HeaderKeywords = new string[14]
	{
		"id", "name", "date", "amount", "total", "price", "quantity", "日期", "名称", "金额",
		"数量", "编号", "类型", "状态"
	};

	public static int Run(string[] args)
	{
		if (args.Length < 1)
		{
			Console.WriteLine("Usage: KimiXlsx inspect <input.xlsx> [--output/-o <file.json>] [--pretty/-p]");
			return 1;
		}
		string filePath = args[0];
		string text = null;
		bool writeIndented = false;
		for (int i = 1; i < args.Length; i++)
		{
			if ((args[i] == "--output" || args[i] == "-o") && i + 1 < args.Length)
			{
				text = args[++i];
			}
			else if (args[i] == "--pretty" || args[i] == "-p")
			{
				writeIndented = true;
			}
		}
		InspectionResult inspectionResult = InspectExcel(filePath);
		JsonSerializerOptions options = new JsonSerializerOptions
		{
			WriteIndented = writeIndented,
			PropertyNamingPolicy = JsonNamingPolicy.CamelCase
		};
		string text2 = JsonSerializer.Serialize(inspectionResult, options);
		if (text != null)
		{
			File.WriteAllText(text, text2);
			Console.WriteLine("Structure saved to: " + text);
		}
		else
		{
			Console.WriteLine(text2);
		}
		return (!(inspectionResult.Status == "success")) ? 1 : 0;
	}

	public static string IndexToColumn(int idx)
	{
		string text = "";
		for (idx++; idx > 0; idx /= 26)
		{
			idx--;
			text = (char)(65 + idx % 26) + text;
		}
		return text;
	}

	public static int ColumnToIndex(string col)
	{
		int num = 0;
		string text = col.ToUpper();
		foreach (char c in text)
		{
			num = num * 26 + (c - 65 + 1);
		}
		return num - 1;
	}

	public static string MakeRange(int startCol, int startRow, int endCol, int endRow)
	{
		return $"{IndexToColumn(startCol)}{startRow}:{IndexToColumn(endCol)}{endRow}";
	}

	public static InspectionResult InspectExcel(string filePath)
	{
		if (!File.Exists(filePath))
		{
			return new InspectionResult
			{
				Status = "error",
				Message = "File not found: " + filePath
			};
		}
		try
		{
			using XLWorkbook xLWorkbook = new XLWorkbook(filePath);
			InspectionResult inspectionResult = new InspectionResult
			{
				File = Path.GetFileName(filePath)
			};
			foreach (IXLWorksheet worksheet in xLWorkbook.Worksheets)
			{
				SheetInfo item = AnalyzeSheet(worksheet);
				inspectionResult.Sheets.Add(item);
			}
			return inspectionResult;
		}
		catch (Exception ex)
		{
			return new InspectionResult
			{
				Status = "error",
				Message = ex.Message
			};
		}
	}

	private static SheetInfo AnalyzeSheet(IXLWorksheet ws)
	{
		SheetInfo sheetInfo = new SheetInfo
		{
			Name = ws.Name
		};
		IXLCell iXLCell = ws.FirstCellUsed();
		IXLCell iXLCell2 = ws.LastCellUsed();
		if (iXLCell == null || iXLCell2 == null)
		{
			sheetInfo.TotalRows = 0;
			sheetInfo.TotalCols = 0;
			return sheetInfo;
		}
		int rowNumber = iXLCell.Address.RowNumber;
		int rowNumber2 = iXLCell2.Address.RowNumber;
		int columnNumber = iXLCell.Address.ColumnNumber;
		int columnNumber2 = iXLCell2.Address.ColumnNumber;
		sheetInfo.TotalRows = rowNumber2 - rowNumber + 1;
		sheetInfo.TotalCols = columnNumber2 - columnNumber + 1;
		sheetInfo.DataRange = MakeRange(columnNumber - 1, rowNumber, columnNumber2 - 1, rowNumber2);
		Dictionary<(int, int), object> dictionary = new Dictionary<(int, int), object>();
		foreach (IXLCell item2 in ws.CellsUsed())
		{
			int rowNumber3 = item2.Address.RowNumber;
			int item = item2.Address.ColumnNumber - 1;
			object value = ((!item2.HasFormula) ? ((!item2.Value.IsNumber) ? ((!item2.Value.IsBoolean) ? ((!item2.Value.IsText) ? item2.Value.ToString() : item2.Value.GetText()) : ((object)item2.Value.GetBoolean())) : ((object)item2.Value.GetNumber())) : new
			{
				formula = item2.FormulaA1
			});
			dictionary[(rowNumber3, item)] = value;
		}
		List<TableInfo> tables = DetectTables(dictionary, rowNumber, rowNumber2, columnNumber - 1, columnNumber2 - 1);
		sheetInfo.Tables = tables;
		return sheetInfo;
	}

	private static List<TableInfo> DetectTables(Dictionary<(int, int), object> cellsData, int minRow, int maxRow, int minCol, int maxCol)
	{
		List<TableInfo> list = new List<TableInfo>();
		HashSet<(int, int)> hashSet = new HashSet<(int, int)>();
		for (int i = minRow; i <= maxRow; i++)
		{
			for (int j = minCol; j <= maxCol; j++)
			{
				if (hashSet.Contains((i, j)) || !cellsData.ContainsKey((i, j)))
				{
					continue;
				}
				(int, int, int, int)? tuple = ExpandTableRegion(cellsData, i, j, minRow, maxRow, minCol, maxCol);
				if (!tuple.HasValue)
				{
					continue;
				}
				(int, int, int, int) value = tuple.Value;
				int item = value.Item1;
				int item2 = value.Item2;
				int item3 = value.Item3;
				int item4 = value.Item4;
				for (int k = item; k <= item2; k++)
				{
					for (int l = item3; l <= item4; l++)
					{
						hashSet.Add((k, l));
					}
				}
				TableInfo item5 = AnalyzeTableRegion(cellsData, item, item2, item3, item4);
				list.Add(item5);
			}
		}
		return list;
	}

	private static (int, int, int, int)? ExpandTableRegion(Dictionary<(int, int), object> cellsData, int startRow, int startCol, int minRow, int maxRow, int minCol, int maxCol)
	{
		int num = startRow;
		int num2 = startRow;
		int num3 = startCol;
		int num4 = startCol;
		for (int i = startCol; i <= maxCol; i++)
		{
			if (cellsData.ContainsKey((startRow, i)))
			{
				num4 = i;
			}
			else if (i + 1 > maxCol || !cellsData.ContainsKey((startRow, i + 1)))
			{
				break;
			}
		}
		int num5 = startCol - 1;
		while (num5 >= minCol && cellsData.ContainsKey((startRow, num5)))
		{
			num3 = num5;
			num5--;
		}
		int row;
		for (row = startRow + 1; row <= maxRow; row++)
		{
			bool flag = false;
			for (int j = num3; j <= num4; j++)
			{
				if (cellsData.ContainsKey((row, j)))
				{
					flag = true;
					break;
				}
			}
			if (flag)
			{
				num2 = row;
			}
			else if (row + 1 > maxRow || !Enumerable.Range(num3, num4 - num3 + 1).Any((int c) => cellsData.ContainsKey((row + 1, c))))
			{
				break;
			}
		}
		for (int num6 = startRow - 1; num6 >= minRow; num6--)
		{
			bool flag2 = false;
			for (int num7 = num3; num7 <= num4; num7++)
			{
				if (cellsData.ContainsKey((num6, num7)))
				{
					flag2 = true;
					break;
				}
			}
			if (!flag2)
			{
				break;
			}
			num = num6;
		}
		if (num2 - num < 1 || num4 - num3 < 0)
		{
			return null;
		}
		return (num, num2, num3, num4);
	}

	private static TableInfo AnalyzeTableRegion(Dictionary<(int, int), object> cellsData, int minRow, int maxRow, int minCol, int maxCol)
	{
		List<object> list = new List<object>();
		for (int i = minCol; i <= maxCol; i++)
		{
			list.Add(cellsData.GetValueOrDefault((minRow, i)));
		}
		bool flag = LooksLikeHeader(list, cellsData, minRow, maxRow, minCol, maxCol);
		int? totalRow = null;
		for (int num = maxRow; num > minRow; num--)
		{
			object valueOrDefault = cellsData.GetValueOrDefault((num, minCol));
			string str = valueOrDefault as string;
			if (str != null && TotalKeywords.Any((string kw) => str.ToLower().Contains(kw)))
			{
				totalRow = num;
				break;
			}
		}
		string range = MakeRange(minCol, minRow, maxCol, maxRow);
		TableInfo tableInfo = new TableInfo
		{
			Range = range,
			StartRow = minRow,
			EndRow = maxRow,
			StartCol = IndexToColumn(minCol),
			EndCol = IndexToColumn(maxCol),
			RowCount = maxRow - minRow + 1,
			ColCount = maxCol - minCol + 1
		};
		if (flag)
		{
			tableInfo.HeaderRow = minRow;
			tableInfo.Headers = list.Select((object h) => h?.ToString() ?? "").ToList();
			tableInfo.DataStartRow = minRow + 1;
		}
		else
		{
			tableInfo.HeaderRow = null;
			tableInfo.DataStartRow = minRow;
		}
		if (totalRow.HasValue)
		{
			tableInfo.TotalRow = totalRow;
			tableInfo.DataEndRow = totalRow.Value - 1;
		}
		else
		{
			tableInfo.TotalRow = null;
			tableInfo.DataEndRow = maxRow;
		}
		return tableInfo;
	}

	private static bool LooksLikeHeader(List<object?> headerRow, Dictionary<(int, int), object> cellsData, int minRow, int maxRow, int minCol, int maxCol)
	{
		if (headerRow.Count == 0)
		{
			return false;
		}
		bool flag = headerRow.Where((object v) => v != null).All((object v) => v is string);
		if (maxRow > minRow)
		{
			bool flag2 = (from v in (from c in Enumerable.Range(minCol, maxCol - minCol + 1)
					select cellsData.GetValueOrDefault((minRow + 1, c))).ToList()
				where v != null
				select v).Any((object v) => (v is double || v is int || v is long) ? true : false);
			if (flag && flag2)
			{
				return true;
			}
		}
		foreach (object item in headerRow)
		{
			string str = item as string;
			if (str != null && HeaderKeywords.Any((string kw) => str.ToLower().Contains(kw)))
			{
				return true;
			}
		}
		if (flag)
		{
			return headerRow.Count((object v) => v != null) >= 2;
		}
		return false;
	}
}
