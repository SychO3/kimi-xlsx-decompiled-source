using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace KimiXlsx.Commands;

public static class ReferenceCheckCommand
{
	private static readonly Regex CellReferencePattern = new Regex("\\$?[A-Z]+\\$?\\d+(?::\\$?[A-Z]+\\$?\\d+)?", RegexOptions.Compiled);

	private static readonly Regex RowNumberPattern = new Regex("\\d+", RegexOptions.Compiled);

	private static readonly Regex HeaderReferencePattern = new Regex("[A-Z]+1(?:[^0-9]|$)", RegexOptions.Compiled);

	private static readonly Regex AggregatePattern = new Regex("(SUM|AVERAGE|COUNT)\\s*\\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly Regex RangePattern = new Regex("([A-Z]+)(\\d+):([A-Z]+)(\\d+)", RegexOptions.Compiled);

	private static readonly Regex CellRefNormalize = new Regex("([A-Z]+)(\\d+)", RegexOptions.Compiled);

	public static int Run(string[] args)
	{
		if (args.Length < 1)
		{
			Console.WriteLine("[Help] KimiXlsx reference-check <excel_file_path>");
			return 1;
		}
		ValidationResult validationResult = ValidateFormulas(args[0]);
		Console.Write(FormatYamlOutput(validationResult));
		return (validationResult.Error != null) ? 1 : 0;
	}

	public static List<string> ExtractCellReferences(string formula)
	{
		return (from m in CellReferencePattern.Matches(formula)
			select m.Value).ToList();
	}

	public static string NormalizeFormula(string formula, int row, int col)
	{
		return CellRefNormalize.Replace(formula, delegate(Match match)
		{
			string value = match.Groups[1].Value;
			int num = int.Parse(match.Groups[2].Value) - row;
			return $"{value}{{{((num >= 0) ? "+" : "")}{num}}}";
		});
	}

	public static List<Issue> AnalyzeFormulaReferences(IXLWorksheet ws, string sheetName, int maxDataRow, int maxDataCol)
	{
		List<Issue> list = new List<Issue>();
		foreach (IXLCell item in ws.CellsUsed())
		{
			if (!item.HasFormula)
			{
				continue;
			}
			string formulaA = item.FormulaA1;
			string location = $"{sheetName}!{item.Address}";
			foreach (string item2 in ExtractCellReferences(formulaA))
			{
				if (!item2.Contains(':'))
				{
					continue;
				}
				string[] array = item2.Split(':');
				Match match = RowNumberPattern.Match(array[1]);
				if (match.Success)
				{
					int num = int.Parse(match.Value);
					if ((double)num > (double)maxDataRow * 1.5)
					{
						list.Add(new Issue
						{
							Type = "oversized_range",
							Location = location,
							Formula = formulaA,
							IssueDescription = $"Reference range extends to row {num}, but data only goes to row {maxDataRow}"
						});
					}
				}
			}
			if (HeaderReferencePattern.IsMatch(formulaA))
			{
				list.Add(new Issue
				{
					Type = "header_reference",
					Location = location,
					Formula = formulaA,
					IssueDescription = "Possible reference to header row (row 1)"
				});
			}
			if (!AggregatePattern.IsMatch(formulaA))
			{
				continue;
			}
			Match match2 = RangePattern.Match(formulaA);
			if (match2.Success)
			{
				int num2 = int.Parse(match2.Groups[2].Value);
				int num3 = int.Parse(match2.Groups[4].Value) - num2 + 1;
				if (num3 <= 2)
				{
					list.Add(new Issue
					{
						Type = "small_aggregate_range",
						Location = location,
						Formula = formulaA,
						IssueDescription = $"Aggregate function only covers {num3} cells, data may be missing"
					});
				}
			}
		}
		return list;
	}

	public static List<Issue> CheckFormulaConsistency(IXLWorksheet ws, string sheetName)
	{
		List<Issue> list = new List<Issue>();
		Dictionary<string, List<(string Cell, string Formula, string Normalized, int Row)>> dictionary = new Dictionary<string, List<(string Cell, string Formula, string Normalized, int Row)>>();
		foreach (IXLCell item2 in ws.CellsUsed())
		{
			if (item2.HasFormula)
			{
				string columnLetter = item2.Address.ColumnLetter;
				string item = NormalizeFormula(item2.FormulaA1, item2.Address.RowNumber, item2.Address.ColumnNumber);
				if (!dictionary.ContainsKey(columnLetter))
				{
					dictionary[columnLetter] = new List<(string Cell, string Formula, string Normalized, int Row)>();
				}
				dictionary[columnLetter].Add(($"{sheetName}!{item2.Address}", item2.FormulaA1, item, item2.Address.RowNumber));
			}
		}
		foreach (var (_, list3) in dictionary)
		{
			if (list3.Count < 3)
			{
				continue;
			}
			KeyValuePair<string, int> keyValuePair2 = (from kv in (from f in list3
					group f by f.Normalized).ToDictionary((IGrouping<string, (string Cell, string Formula, string Normalized, int Row)> g) => g.Key, (IGrouping<string, (string Cell, string Formula, string Normalized, int Row)> g) => g.Count())
				orderby kv.Value descending
				select kv).First();
			foreach (var item3 in list3)
			{
				if (item3.Item3 != keyValuePair2.Key)
				{
					list.Add(new Issue
					{
						Type = "inconsistent_formula",
						Location = item3.Item1,
						Formula = item3.Item2,
						IssueDescription = $"This column has {keyValuePair2.Value} cells using a different formula pattern",
						ExpectedPattern = keyValuePair2.Key,
						ActualPattern = item3.Item3
					});
				}
			}
		}
		return list;
	}

	public static ValidationResult ValidateFormulas(string filename)
	{
		try
		{
			using XLWorkbook xLWorkbook = new XLWorkbook(filename);
			List<Issue> list = new List<Issue>();
			Dictionary<string, int> dictionary = new Dictionary<string, int>
			{
				["oversized_range"] = 0,
				["header_reference"] = 0,
				["small_aggregate_range"] = 0,
				["inconsistent_formula"] = 0
			};
			foreach (IXLWorksheet worksheet in xLWorkbook.Worksheets)
			{
				string name = worksheet.Name;
				int maxDataRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
				int maxDataCol = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
				List<Issue> collection = AnalyzeFormulaReferences(worksheet, name, maxDataRow, maxDataCol);
				list.AddRange(collection);
				List<Issue> collection2 = CheckFormulaConsistency(worksheet, name);
				list.AddRange(collection2);
			}
			foreach (Issue item in list)
			{
				if (dictionary.ContainsKey(item.Type))
				{
					dictionary[item.Type]++;
				}
			}
			return new ValidationResult
			{
				TotalIssues = list.Count,
				IssuesByType = dictionary,
				Issues = list
			};
		}
		catch (Exception ex)
		{
			return new ValidationResult
			{
				Error = ex.Message
			};
		}
	}

	public static string FormatYamlOutput(ValidationResult result)
	{
		StringBuilder stringBuilder = new StringBuilder();
		StringBuilder stringBuilder2;
		StringBuilder.AppendInterpolatedStringHandler handler;
		if (result.Error != null)
		{
			stringBuilder2 = stringBuilder;
			StringBuilder stringBuilder3 = stringBuilder2;
			handler = new StringBuilder.AppendInterpolatedStringHandler(7, 1, stringBuilder2);
			handler.AppendLiteral("error: ");
			handler.AppendFormatted(result.Error);
			stringBuilder3.AppendLine(ref handler);
			return stringBuilder.ToString();
		}
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder4 = stringBuilder2;
		handler = new StringBuilder.AppendInterpolatedStringHandler(8, 1, stringBuilder2);
		handler.AppendLiteral("result: ");
		handler.AppendFormatted((result.TotalIssues == 0) ? "pass" : "has_issues");
		stringBuilder4.AppendLine(ref handler);
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder5 = stringBuilder2;
		handler = new StringBuilder.AppendInterpolatedStringHandler(14, 1, stringBuilder2);
		handler.AppendLiteral("total_issues: ");
		handler.AppendFormatted(result.TotalIssues);
		stringBuilder5.AppendLine(ref handler);
		stringBuilder.AppendLine("issues_by_type:");
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder6 = stringBuilder2;
		handler = new StringBuilder.AppendInterpolatedStringHandler(19, 1, stringBuilder2);
		handler.AppendLiteral("  oversized_range: ");
		handler.AppendFormatted(result.IssuesByType.GetValueOrDefault("oversized_range", 0));
		stringBuilder6.AppendLine(ref handler);
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder7 = stringBuilder2;
		handler = new StringBuilder.AppendInterpolatedStringHandler(20, 1, stringBuilder2);
		handler.AppendLiteral("  header_reference: ");
		handler.AppendFormatted(result.IssuesByType.GetValueOrDefault("header_reference", 0));
		stringBuilder7.AppendLine(ref handler);
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder8 = stringBuilder2;
		handler = new StringBuilder.AppendInterpolatedStringHandler(25, 1, stringBuilder2);
		handler.AppendLiteral("  small_aggregate_range: ");
		handler.AppendFormatted(result.IssuesByType.GetValueOrDefault("small_aggregate_range", 0));
		stringBuilder8.AppendLine(ref handler);
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder9 = stringBuilder2;
		handler = new StringBuilder.AppendInterpolatedStringHandler(24, 1, stringBuilder2);
		handler.AppendLiteral("  inconsistent_formula: ");
		handler.AppendFormatted(result.IssuesByType.GetValueOrDefault("inconsistent_formula", 0));
		stringBuilder9.AppendLine(ref handler);
		if (result.Issues.Count > 0)
		{
			stringBuilder.AppendLine("issues:");
			foreach (Issue item in result.Issues.Take(20))
			{
				stringBuilder2 = stringBuilder;
				StringBuilder stringBuilder10 = stringBuilder2;
				handler = new StringBuilder.AppendInterpolatedStringHandler(10, 1, stringBuilder2);
				handler.AppendLiteral("  - type: ");
				handler.AppendFormatted(item.Type);
				stringBuilder10.AppendLine(ref handler);
				stringBuilder2 = stringBuilder;
				StringBuilder stringBuilder11 = stringBuilder2;
				handler = new StringBuilder.AppendInterpolatedStringHandler(14, 1, stringBuilder2);
				handler.AppendLiteral("    location: ");
				handler.AppendFormatted(item.Location);
				stringBuilder11.AppendLine(ref handler);
				stringBuilder2 = stringBuilder;
				StringBuilder stringBuilder12 = stringBuilder2;
				handler = new StringBuilder.AppendInterpolatedStringHandler(13, 1, stringBuilder2);
				handler.AppendLiteral("    formula: ");
				handler.AppendFormatted(item.Formula);
				stringBuilder12.AppendLine(ref handler);
				stringBuilder2 = stringBuilder;
				StringBuilder stringBuilder13 = stringBuilder2;
				handler = new StringBuilder.AppendInterpolatedStringHandler(11, 1, stringBuilder2);
				handler.AppendLiteral("    issue: ");
				handler.AppendFormatted(item.IssueDescription);
				stringBuilder13.AppendLine(ref handler);
				if (item.ExpectedPattern != null)
				{
					stringBuilder2 = stringBuilder;
					StringBuilder stringBuilder14 = stringBuilder2;
					handler = new StringBuilder.AppendInterpolatedStringHandler(22, 1, stringBuilder2);
					handler.AppendLiteral("    expected_pattern: ");
					handler.AppendFormatted(item.ExpectedPattern);
					stringBuilder14.AppendLine(ref handler);
				}
				if (item.ActualPattern != null)
				{
					stringBuilder2 = stringBuilder;
					StringBuilder stringBuilder15 = stringBuilder2;
					handler = new StringBuilder.AppendInterpolatedStringHandler(20, 1, stringBuilder2);
					handler.AppendLiteral("    actual_pattern: ");
					handler.AppendFormatted(item.ActualPattern);
					stringBuilder15.AppendLine(ref handler);
				}
			}
			if (result.Issues.Count > 20)
			{
				stringBuilder2 = stringBuilder;
				StringBuilder stringBuilder16 = stringBuilder2;
				handler = new StringBuilder.AppendInterpolatedStringHandler(24, 1, stringBuilder2);
				handler.AppendLiteral("  # ... and ");
				handler.AppendFormatted(result.Issues.Count - 20);
				handler.AppendLiteral(" more issues");
				stringBuilder16.AppendLine(ref handler);
			}
		}
		return stringBuilder.ToString();
	}
}
