using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace KimiXlsx.Commands;

public static class ValidateCommand
{
	private static readonly List<string> SafeToIgnorePatterns = new List<string>
	{
		"unexpected child element.*:sz", "unexpected child element.*:color", "unexpected child element.*:name", "unexpected child element.*:family", "unexpected child element.*:scheme", "unexpected child element.*:charset", "unexpected child element.*:b\\b", "unexpected child element.*:i\\b", "unexpected child element.*:u\\b", "unexpected child element.*:strike",
		"unexpected child element.*:outline", "unexpected child element.*:shadow", "unexpected child element.*:condense", "unexpected child element.*:extend", "element 'font' has invalid child element", "The element has invalid child element.*font", "unexpected child element.*:fgColor", "unexpected child element.*:bgColor", "unexpected child element.*:patternFill", "element 'fill' has invalid child element",
		"element 'patternFill' has invalid child element", "unexpected child element.*:left", "unexpected child element.*:right", "unexpected child element.*:top", "unexpected child element.*:bottom", "unexpected child element.*:diagonal", "element 'border' has invalid child element", "element 'alignment' has invalid child element", "element 'numFmt' has invalid child element", "element 'cellStyle' has invalid child element",
		"element 'cellXfs' has invalid child element", "element 'cellStyleXfs' has invalid child element", "element 'sheetView' has invalid child element", "unexpected child element.*:pageMargins", "unexpected child element.*:sheetFormatPr", "unexpected child element.*:sheetView", "attribute is not declared", "The required attribute .* is missing", "styles\\.xml.*unexpected child element"
	};

	private static readonly List<string> NeverIgnorePatterns = new List<string>
	{
		"pivotCache", "pivotTable", "pivotField", "cacheField", "cacheSource", "cacheRecords", "pivotCacheDefinition", "pivotCacheRecords", "rowItems", "colItems",
		"pageFields", "dataFields", "chartSpace", "chart\\.xml", "plotArea", "workbook\\.xml", "The part .* is missing", "corrupted", "malformed"
	};

	private static readonly List<string> IncompatibleFunctions = new List<string>
	{
		"FILTER", "IMPORTRANGE", "IMPORTDATA", "IMPORTHTML", "IMPORTXML", "GOOGLEFINANCE", "GOOGLETRANSLATE", "IMAGE", "SPARKLINE", "QUERY",
		"ARRAYFORMULA", "UNIQUE", "SORT", "SORTBY", "SEQUENCE", "RANDARRAY", "XLOOKUP", "XMATCH", "LET", "LAMBDA",
		"MAP", "REDUCE", "SCAN", "MAKEARRAY", "BYROW", "BYCOL", "CHOOSECOLS", "CHOOSEROWS", "DROP", "TAKE",
		"EXPAND", "WRAPCOLS", "WRAPROWS", "TOCOL", "TOROW", "VSTACK", "HSTACK", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER",
		"VALUETOTEXT", "ARRAYTOTEXT"
	};

	public static int Run(string[] args)
	{
		if (args.Length < 1)
		{
			Console.WriteLine("\nKimiXlsx validate - OpenXML Structure Validator\n================================================\n\nUsage: dotnet run --project KimiXlsx -- validate <file.xlsx> [options]\n\nOptions:\n  --json           Output results as JSON\n  --strict         Fail on ANY error (no whitelist filtering)\n  --lenient        Show ignored errors but still pass if only safe issues\n  --delete-invalid Delete file if validation fails\n  --quiet          Only output errors, no success message\n\nDescription:\n  Validates Excel files against the OpenXML specification (Office 2013).\n\n  By default, known safe openpyxl schema issues are filtered out (whitelist),\n  but PivotTable and Chart structure errors are NEVER ignored.\n\n  Use --strict to disable whitelist filtering and fail on any error.\n  Use --lenient to see what errors were ignored.\n\nExit Codes:\n  0 = Validation passed (or only safe ignored errors)\n  1 = Validation failed (critical errors found)\n  2 = File not found or cannot be opened\n");
			return 1;
		}
		string text = args[0];
		bool num = args.Contains("--json");
		bool strict = args.Contains("--strict");
		bool lenient = args.Contains("--lenient");
		bool flag = args.Contains("--delete-invalid");
		bool quiet = args.Contains("--quiet");
		OpenXmlValidationResult openXmlValidationResult = ValidateOpenXml(text, strict);
		if (num)
		{
			JsonSerializerOptions options = new JsonSerializerOptions
			{
				WriteIndented = true
			};
			Console.WriteLine(JsonSerializer.Serialize(openXmlValidationResult, options));
		}
		else
		{
			PrintReport(openXmlValidationResult, quiet, lenient);
		}
		if (openXmlValidationResult.Status == "pass" || openXmlValidationResult.Status == "pass_with_warnings")
		{
			return 0;
		}
		if (flag && File.Exists(text))
		{
			try
			{
				File.Delete(text);
				Console.WriteLine("\n⛔ DELETED invalid file: " + text);
			}
			catch (Exception ex)
			{
				Console.WriteLine("\n⚠\ufe0f  Could not delete file: " + ex.Message);
			}
		}
		if (openXmlValidationResult.FatalError == null)
		{
			return 1;
		}
		return 2;
	}

	private static bool MatchesAnyPattern(string description, string? path, string? part, List<string> patterns)
	{
		string input = $"{description} {path ?? ""} {part ?? ""}";
		foreach (string pattern in patterns)
		{
			if (Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase))
			{
				return true;
			}
		}
		return false;
	}

	private static (bool isCritical, bool isSafeToIgnore, string? reason) ClassifyError(string description, string? path, string? part)
	{
		if (MatchesAnyPattern(description, path, part, NeverIgnorePatterns))
		{
			return (isCritical: true, isSafeToIgnore: false, reason: null);
		}
		if (MatchesAnyPattern(description, path, part, SafeToIgnorePatterns))
		{
			return (isCritical: false, isSafeToIgnore: true, reason: "Known openpyxl schema ordering issue");
		}
		return (isCritical: false, isSafeToIgnore: false, reason: null);
	}

	private static List<(string RelsFile, string Target, string RelationshipId)> CheckAbsolutePathsInRels(string filePath)
	{
		List<(string, string, string)> list = new List<(string, string, string)>();
		try
		{
			using ZipArchive zipArchive = ZipFile.OpenRead(filePath);
			foreach (ZipArchiveEntry item2 in zipArchive.Entries.Where((ZipArchiveEntry e) => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToList())
			{
				using Stream stream = item2.Open();
				XDocument xDocument = XDocument.Load(stream);
				XNamespace xNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
				foreach (XElement item3 in xDocument.Descendants(xNamespace + "Relationship"))
				{
					string text = item3.Attribute("Target")?.Value;
					string item = item3.Attribute("Id")?.Value ?? "?";
					if (text != null && text.StartsWith("/"))
					{
						list.Add((item2.FullName, text, item));
						if (list.Count >= 20)
						{
							return list;
						}
					}
				}
			}
		}
		catch
		{
		}
		return list;
	}

	private static List<(string ChartName, string IssueType, string Details)> CheckChartStructureIssues(string filePath)
	{
		List<(string, string, string)> list = new List<(string, string, string)>();
		try
		{
			using ZipArchive zipArchive = ZipFile.OpenRead(filePath);
			foreach (ZipArchiveEntry item in zipArchive.Entries.Where((ZipArchiveEntry e) => e.FullName.Contains("charts/chart") && e.FullName.EndsWith(".xml")).ToList())
			{
				using Stream stream = item.Open();
				XDocument xDocument = XDocument.Load(stream);
				XNamespace xNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
				string name = item.Name;
				foreach (XElement item2 in xDocument.Descendants(xNamespace + "ser").ToList())
				{
					string value = item2.Element(xNamespace + "idx")?.Attribute("val")?.Value ?? "?";
					XElement xElement = item2.Element(xNamespace + "cat");
					int result = 0;
					if (xElement != null)
					{
						XElement? xElement2 = xElement.Element(xNamespace + "strRef");
						XElement xElement3 = xElement.Element(xNamespace + "numRef");
						XElement xElement4 = xElement2?.Element(xNamespace + "strCache") ?? xElement3?.Element(xNamespace + "numCache");
						if (xElement4 != null)
						{
							XElement xElement5 = xElement4.Element(xNamespace + "ptCount");
							if (xElement5 != null)
							{
								int.TryParse(xElement5.Attribute("val")?.Value, out result);
							}
							List<XElement> list2 = xElement4.Elements(xNamespace + "pt").ToList();
							_ = list2.Count;
							foreach (XElement item3 in list2)
							{
								XAttribute xAttribute = item3.Attribute("idx");
								if (xAttribute != null && int.TryParse(xAttribute.Value, out var result2) && result2 >= result && result > 0)
								{
									list.Add((name, "CategoryPointIndexOutOfBounds", $"Series {value}: category point idx={result2} exceeds declared ptCount={result}"));
								}
							}
						}
					}
					XElement xElement6 = item2.Element(xNamespace + "val");
					int result3 = 0;
					if (xElement6 != null)
					{
						XElement xElement7 = xElement6.Element(xNamespace + "numRef")?.Element(xNamespace + "numCache");
						if (xElement7 != null)
						{
							XElement xElement8 = xElement7.Element(xNamespace + "ptCount");
							if (xElement8 != null)
							{
								int.TryParse(xElement8.Attribute("val")?.Value, out result3);
							}
							List<XElement> list3 = xElement7.Elements(xNamespace + "pt").ToList();
							_ = list3.Count;
							foreach (XElement item4 in list3)
							{
								XAttribute xAttribute2 = item4.Attribute("idx");
								if (xAttribute2 != null && int.TryParse(xAttribute2.Value, out var result4) && result4 >= result3 && result3 > 0)
								{
									list.Add((name, "ValuePointIndexOutOfBounds", $"Series {value}: value point idx={result4} exceeds declared ptCount={result3}"));
								}
							}
						}
					}
					if (result > 0 && result3 > 0 && result != result3)
					{
						list.Add((name, "CategoryValueCountMismatch", $"Series {value}: category has {result} points but values has {result3} points"));
					}
				}
				if (list.Count >= 20)
				{
					break;
				}
			}
		}
		catch (Exception ex)
		{
			list.Add(("Unknown", "ValidationError", "Error checking Chart structure: " + ex.Message));
		}
		return list;
	}

	private static List<(string PivotTableName, string IssueType, string Details)> CheckPivotTableIndexBounds(string filePath)
	{
		List<(string, string, string)> list = new List<(string, string, string)>();
		try
		{
			using ZipArchive zipArchive = ZipFile.OpenRead(filePath);
			Dictionary<string, List<int>> dictionary = new Dictionary<string, List<int>>();
			foreach (ZipArchiveEntry item3 in zipArchive.Entries.Where((ZipArchiveEntry e) => e.FullName.Contains("pivotCache/pivotCacheDefinition") && e.FullName.EndsWith(".xml")).ToList())
			{
				using Stream stream = item3.Open();
				XDocument xDocument = XDocument.Load(stream);
				XNamespace xNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
				List<XElement> list2 = xDocument.Descendants(xNamespace + "cacheField").ToList();
				List<int> list3 = new List<int>();
				foreach (XElement item4 in list2)
				{
					XElement xElement = item4.Element(xNamespace + "sharedItems");
					if (xElement != null)
					{
						XAttribute xAttribute = xElement.Attribute("count");
						if (xAttribute != null && int.TryParse(xAttribute.Value, out var result))
						{
							list3.Add(result);
							continue;
						}
						int item = xElement.Elements().Count();
						list3.Add(item);
					}
					else
					{
						list3.Add(0);
					}
				}
				Match match = Regex.Match(item3.FullName, "pivotCacheDefinition(\\d+)\\.xml");
				if (match.Success)
				{
					dictionary[match.Groups[1].Value] = list3;
				}
			}
			foreach (ZipArchiveEntry item5 in zipArchive.Entries.Where((ZipArchiveEntry e) => e.FullName.Contains("pivotTables/pivotTable") && e.FullName.EndsWith(".xml")).ToList())
			{
				using Stream stream2 = item5.Open();
				XDocument xDocument2 = XDocument.Load(stream2);
				XNamespace xNamespace2 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
				XElement root = xDocument2.Root;
				if (root == null)
				{
					continue;
				}
				string item2 = root.Attribute("name")?.Value ?? item5.Name;
				_ = root.Attribute("cacheId")?.Value;
				List<XElement> list4 = xDocument2.Descendants(xNamespace2 + "pivotField").ToList();
				Dictionary<int, int> dictionary2 = new Dictionary<int, int>();
				for (int num = 0; num < list4.Count; num++)
				{
					XElement xElement2 = list4[num].Element(xNamespace2 + "items");
					if (xElement2 != null)
					{
						XAttribute xAttribute2 = xElement2.Attribute("count");
						if (xAttribute2 != null && int.TryParse(xAttribute2.Value, out var result2))
						{
							dictionary2[num] = result2 - 1;
						}
					}
				}
				XElement xElement3 = xDocument2.Descendants(xNamespace2 + "rowItems").FirstOrDefault();
				if (xElement3 != null)
				{
					string text = xElement3.Attribute("count")?.Value;
					List<XElement> list5 = xElement3.Elements(xNamespace2 + "i").ToList();
					XElement xElement4 = xDocument2.Descendants(xNamespace2 + "rowFields").FirstOrDefault();
					int num2 = -1;
					if (xElement4 != null)
					{
						XElement xElement5 = xElement4.Elements(xNamespace2 + "field").FirstOrDefault();
						if (xElement5 != null)
						{
							XAttribute xAttribute3 = xElement5.Attribute("x");
							if (xAttribute3 != null && int.TryParse(xAttribute3.Value, out var result3))
							{
								num2 = result3;
							}
						}
					}
					foreach (XElement item6 in list5)
					{
						XElement xElement6 = item6.Element(xNamespace2 + "x");
						if (xElement6 != null)
						{
							XAttribute xAttribute4 = xElement6.Attribute("v");
							if (xAttribute4 != null && int.TryParse(xAttribute4.Value, out var result4) && num2 >= 0 && dictionary2.ContainsKey(num2) && result4 > dictionary2[num2])
							{
								list.Add((item2, "RowItemIndexOutOfBounds", $"rowItem references index {result4}, but field[{num2}] only has indices 0-{dictionary2[num2]}"));
							}
						}
					}
					if (text != null && int.TryParse(text, out var result5) && num2 >= 0 && dictionary2.ContainsKey(num2))
					{
						int num3 = dictionary2[num2] + 2;
						if (result5 > num3)
						{
							list.Add((item2, "RowItemsCountMismatch", $"rowItems declares {result5} items, but field[{num2}] only supports up to {num3} items"));
						}
					}
				}
				XElement xElement7 = xDocument2.Descendants(xNamespace2 + "colItems").FirstOrDefault();
				if (xElement7 != null)
				{
					XElement xElement8 = xDocument2.Descendants(xNamespace2 + "colFields").FirstOrDefault();
					int num4 = -1;
					if (xElement8 != null)
					{
						XElement xElement9 = xElement8.Elements(xNamespace2 + "field").FirstOrDefault();
						if (xElement9 != null)
						{
							XAttribute xAttribute5 = xElement9.Attribute("x");
							if (xAttribute5 != null && int.TryParse(xAttribute5.Value, out var result6))
							{
								num4 = result6;
							}
						}
					}
					foreach (XElement item7 in xElement7.Elements(xNamespace2 + "i").ToList())
					{
						XElement xElement10 = item7.Element(xNamespace2 + "x");
						if (xElement10 != null)
						{
							XAttribute xAttribute6 = xElement10.Attribute("v");
							if (xAttribute6 != null && int.TryParse(xAttribute6.Value, out var result7) && num4 >= 0 && dictionary2.ContainsKey(num4) && result7 > dictionary2[num4])
							{
								list.Add((item2, "ColItemIndexOutOfBounds", $"colItem references index {result7}, but field[{num4}] only has indices 0-{dictionary2[num4]}"));
							}
						}
					}
				}
				if (list.Count < 20)
				{
					continue;
				}
				break;
			}
		}
		catch (Exception ex)
		{
			list.Add(("Unknown", "ValidationError", "Error checking PivotTable indices: " + ex.Message));
		}
		return list;
	}

	private static List<(string SheetName, string CellRef, string Function, string Formula)> CheckIncompatibleFormulas(SpreadsheetDocument doc)
	{
		List<(string, string, string, string)> list = new List<(string, string, string, string)>();
		WorkbookPart workbookPart = doc.WorkbookPart;
		if (workbookPart == null)
		{
			return list;
		}
		Sheets sheets = workbookPart.Workbook.Sheets;
		if (sheets == null)
		{
			return list;
		}
		foreach (Sheet item4 in sheets.Elements<Sheet>())
		{
			string item = item4.Name?.Value ?? "Unknown";
			string text = item4.Id?.Value;
			if (text == null || !(workbookPart.GetPartById(text) is WorksheetPart worksheetPart))
			{
				continue;
			}
			SheetData firstChild = worksheetPart.Worksheet.GetFirstChild<SheetData>();
			if (firstChild == null)
			{
				continue;
			}
			foreach (Row item5 in firstChild.Elements<Row>())
			{
				foreach (Cell item6 in item5.Elements<Cell>())
				{
					string text2 = item6.CellFormula?.Text;
					if (string.IsNullOrEmpty(text2))
					{
						continue;
					}
					string input = text2.ToUpperInvariant();
					foreach (string incompatibleFunction in IncompatibleFunctions)
					{
						string pattern = "\\b" + incompatibleFunction + "\\s*\\(";
						if (Regex.IsMatch(input, pattern, RegexOptions.IgnoreCase))
						{
							string item2 = item6.CellReference?.Value ?? "?";
							string item3 = ((text2.Length > 100) ? (text2.Substring(0, 100) + "...") : text2);
							list.Add((item, item2, incompatibleFunction, item3));
							if (list.Count >= 50)
							{
								return list;
							}
							break;
						}
					}
				}
			}
		}
		return list;
	}

	public static OpenXmlValidationResult ValidateOpenXml(string filePath, bool strict = false)
	{
		OpenXmlValidationResult openXmlValidationResult = new OpenXmlValidationResult
		{
			File = Path.GetFileName(filePath)
		};
		if (!File.Exists(filePath))
		{
			openXmlValidationResult.Status = "fatal_error";
			openXmlValidationResult.FatalError = "File not found: " + filePath;
			return openXmlValidationResult;
		}
		try
		{
			using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, isEditable: false);
			List<ValidationErrorInfo> list = new OpenXmlValidator(FileFormatVersions.Office2013).Validate(spreadsheetDocument).ToList();
			List<ValidationErrorInfo> list2 = (strict ? list : list.Where((ValidationErrorInfo e) => e.ErrorType == ValidationErrorType.Schema || e.ErrorType == ValidationErrorType.Semantic).ToList());
			openXmlValidationResult.TotalErrors = list2.Count;
			List<ValidationIssue> list3 = new List<ValidationIssue>();
			List<ValidationIssue> list4 = new List<ValidationIssue>();
			List<ValidationIssue> list5 = new List<ValidationIssue>();
			foreach (ValidationErrorInfo item18 in list2)
			{
				string description = item18.Description ?? "Unknown error";
				string path = item18.Path?.XPath;
				string part = item18.Part?.Uri?.ToString();
				ValidationIssue validationIssue = new ValidationIssue
				{
					ErrorType = item18.ErrorType.ToString(),
					Description = description,
					Path = path,
					Part = part
				};
				if (strict)
				{
					list3.Add(validationIssue);
					continue;
				}
				var (flag, flag2, ignoreReason) = ClassifyError(description, path, part);
				if (flag)
				{
					list3.Add(validationIssue);
				}
				else if (flag2)
				{
					validationIssue.Ignored = true;
					validationIssue.IgnoreReason = ignoreReason;
					list4.Add(validationIssue);
				}
				else
				{
					list5.Add(validationIssue);
				}
			}
			openXmlValidationResult.CriticalErrors = list3.Count + list5.Count;
			openXmlValidationResult.IgnoredErrors = list4.Count;
			openXmlValidationResult.Errors = list3.Concat(list5).Take(50).ToList();
			openXmlValidationResult.IgnoredIssues = list4.Take(20).ToList();
			List<ValidationIssue> list6 = list3.Concat(list5).ToList();
			if (list6.Count > 0)
			{
				foreach (IGrouping<string, ValidationIssue> item19 in from e in list6
					group e by e.ErrorType into g
					orderby g.Count() descending
					select g)
				{
					openXmlValidationResult.ErrorsByType[item19.Key] = item19.Count();
				}
			}
			List<(string, string, string, string)> list7 = CheckIncompatibleFormulas(spreadsheetDocument);
			if (list7.Count > 0)
			{
				foreach (var item20 in list7)
				{
					string item = item20.Item1;
					string item2 = item20.Item2;
					string item3 = item20.Item3;
					string item4 = item20.Item4;
					ValidationIssue item5 = new ValidationIssue
					{
						ErrorType = "IncompatibleFormula",
						Description = "Function '" + item3 + "' is not supported in older Excel versions (pre-365/2021)",
						Path = item + "!" + item2,
						Part = "Formula: " + item4
					};
					openXmlValidationResult.Errors.Add(item5);
				}
				openXmlValidationResult.CriticalErrors += list7.Count;
				if (!openXmlValidationResult.ErrorsByType.ContainsKey("IncompatibleFormula"))
				{
					openXmlValidationResult.ErrorsByType["IncompatibleFormula"] = 0;
				}
				openXmlValidationResult.ErrorsByType["IncompatibleFormula"] += list7.Count;
			}
			List<(string, string, string)> list8 = CheckAbsolutePathsInRels(filePath);
			if (list8.Count > 0)
			{
				foreach (var item21 in list8)
				{
					string item6 = item21.Item1;
					string item7 = item21.Item2;
					string item8 = item21.Item3;
					ValidationIssue item9 = new ValidationIssue
					{
						ErrorType = "AbsolutePathInRels",
						Description = "Absolute path in .rels file will cause MS Excel to fail opening",
						Path = item6 + " (" + item8 + ")",
						Part = "Target=\"" + item7 + "\" should be relative path (without leading '/')"
					};
					openXmlValidationResult.Errors.Add(item9);
				}
				openXmlValidationResult.CriticalErrors += list8.Count;
				openXmlValidationResult.ErrorsByType["AbsolutePathInRels"] = list8.Count;
			}
			List<(string, string, string)> list9 = CheckPivotTableIndexBounds(filePath);
			if (list9.Count > 0)
			{
				foreach (var item22 in list9)
				{
					string item10 = item22.Item1;
					string item11 = item22.Item2;
					string item12 = item22.Item3;
					ValidationIssue item13 = new ValidationIssue
					{
						ErrorType = "PivotTableIndexOutOfBounds",
						Description = "PivotTable '" + item10 + "': " + item11,
						Path = item10,
						Part = item12
					};
					openXmlValidationResult.Errors.Add(item13);
				}
				openXmlValidationResult.CriticalErrors += list9.Count;
				if (!openXmlValidationResult.ErrorsByType.ContainsKey("PivotTableIndexOutOfBounds"))
				{
					openXmlValidationResult.ErrorsByType["PivotTableIndexOutOfBounds"] = 0;
				}
				openXmlValidationResult.ErrorsByType["PivotTableIndexOutOfBounds"] += list9.Count;
			}
			List<(string, string, string)> list10 = CheckChartStructureIssues(filePath);
			if (list10.Count > 0)
			{
				foreach (var item23 in list10)
				{
					string item14 = item23.Item1;
					string item15 = item23.Item2;
					string item16 = item23.Item3;
					ValidationIssue item17 = new ValidationIssue
					{
						ErrorType = "ChartStructureError",
						Description = "Chart '" + item14 + "': " + item15,
						Path = item14,
						Part = item16
					};
					openXmlValidationResult.Errors.Add(item17);
				}
				openXmlValidationResult.CriticalErrors += list10.Count;
				if (!openXmlValidationResult.ErrorsByType.ContainsKey("ChartStructureError"))
				{
					openXmlValidationResult.ErrorsByType["ChartStructureError"] = 0;
				}
				openXmlValidationResult.ErrorsByType["ChartStructureError"] += list10.Count;
			}
			if (list3.Count > 0 || list5.Count > 0 || list7.Count > 0 || list8.Count > 0 || list9.Count > 0 || list10.Count > 0)
			{
				openXmlValidationResult.Status = "failed";
			}
			else if (list4.Count > 0)
			{
				openXmlValidationResult.Status = "pass_with_warnings";
			}
			else
			{
				openXmlValidationResult.Status = "pass";
			}
		}
		catch (OpenXmlPackageException ex)
		{
			openXmlValidationResult.Status = "fatal_error";
			openXmlValidationResult.FatalError = "Cannot open file as OpenXML package: " + ex.Message;
		}
		catch (InvalidDataException ex2)
		{
			openXmlValidationResult.Status = "fatal_error";
			openXmlValidationResult.FatalError = "File is not a valid ZIP/XLSX archive: " + ex2.Message;
		}
		catch (Exception ex3)
		{
			openXmlValidationResult.Status = "fatal_error";
			openXmlValidationResult.FatalError = "Unexpected error: " + ex3.Message;
		}
		return openXmlValidationResult;
	}

	private static void PrintReport(OpenXmlValidationResult result, bool quiet, bool lenient)
	{
		if (result.FatalError != null)
		{
			Console.WriteLine("\n❌ FATAL ERROR: " + result.FatalError);
			return;
		}
		if (result.Status == "pass")
		{
			if (!quiet)
			{
				Console.WriteLine("\n✅ VALIDATION PASSED: " + result.File);
				Console.WriteLine("   File conforms to OpenXML specification (Office 2013)");
				Console.WriteLine("   Safe to deliver to user.\n");
			}
			return;
		}
		if (result.Status == "pass_with_warnings")
		{
			if (quiet)
			{
				return;
			}
			Console.WriteLine("\n✅ VALIDATION PASSED: " + result.File);
			Console.WriteLine($"   ({result.IgnoredErrors} known safe openpyxl issues were ignored)");
			Console.WriteLine("   Safe to deliver to user.\n");
			if (!lenient || result.IgnoredIssues.Count <= 0)
			{
				return;
			}
			Console.WriteLine("   ℹ\ufe0f  Ignored Issues (known openpyxl schema quirks):");
			foreach (ValidationIssue item in result.IgnoredIssues.Take(10))
			{
				Console.WriteLine("      - " + item.Description);
			}
			if (result.IgnoredErrors > 10)
			{
				Console.WriteLine($"      ... and {result.IgnoredErrors - 10} more ignored issues");
			}
			Console.WriteLine();
			return;
		}
		Console.WriteLine(new string('=', 70));
		Console.WriteLine("❌ OPENXML VALIDATION FAILED");
		Console.WriteLine(new string('=', 70));
		Console.WriteLine("File: " + result.File);
		Console.WriteLine($"Critical Errors: {result.CriticalErrors}");
		if (result.IgnoredErrors > 0)
		{
			Console.WriteLine($"Ignored (safe) Errors: {result.IgnoredErrors}");
		}
		Console.WriteLine();
		if (result.ErrorsByType.Count > 0)
		{
			Console.WriteLine("Errors by Type:");
			foreach (var (value, value2) in result.ErrorsByType)
			{
				Console.WriteLine($"  {value}: {value2}");
			}
			Console.WriteLine();
		}
		Console.WriteLine("Critical Error Details:");
		foreach (ValidationIssue error in result.Errors)
		{
			Console.WriteLine("\n  [" + error.ErrorType + "]");
			Console.WriteLine("  Description: " + error.Description);
			if (!string.IsNullOrEmpty(error.Part))
			{
				Console.WriteLine("  Part: " + error.Part);
			}
			if (!string.IsNullOrEmpty(error.Path))
			{
				Console.WriteLine("  XPath: " + error.Path);
			}
		}
		if (result.CriticalErrors > 50)
		{
			Console.WriteLine($"\n  ... and {result.CriticalErrors - 50} more critical errors");
		}
		if (lenient && result.IgnoredIssues.Count > 0)
		{
			Console.WriteLine("\n--- Ignored Issues (known openpyxl schema quirks) ---");
			foreach (ValidationIssue item2 in result.IgnoredIssues.Take(5))
			{
				Console.WriteLine("  - " + item2.Description);
			}
			if (result.IgnoredErrors > 5)
			{
				Console.WriteLine($"  ... and {result.IgnoredErrors - 5} more ignored issues");
			}
		}
		Console.WriteLine(new string('=', 70));
		Console.WriteLine("⛔ THIS FILE HAS CRITICAL ERRORS");
		Console.WriteLine("   PivotTable/Chart structure errors or unknown issues were found.");
		Console.WriteLine("   You MUST fix these errors before delivering the file.");
		Console.WriteLine(new string('=', 70));
		Console.WriteLine();
	}
}
