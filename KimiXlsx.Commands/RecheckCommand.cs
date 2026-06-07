using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace KimiXlsx.Commands;

public static class RecheckCommand
{
	private static readonly string[] ExcelErrors = new string[7] { "#DIV/0!", "#NAME?", "#VALUE!", "#REF!", "#NULL!", "#NUM!", "#N/A" };

	private static readonly XNamespace MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

	private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

	private const string MacroContent = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!DOCTYPE script:module PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"module.dtd\">\n<script:module xmlns:script=\"http://openoffice.org/2000/script\" script:name=\"Module1\" script:language=\"StarBasic\">\n    Sub RecalculateAndSave()\n      ThisComponent.calculateAll()\n      ThisComponent.store()\n      ThisComponent.close(True)\n    End Sub\n</script:module>";

	private static readonly string[] ImplicitArrayPatterns = new string[3] { "MATCH\\s*\\(\\s*TRUE\\s*\\(\\s*\\)\\s*,", "MATCH\\s*\\([^,]+[<>=]+[^,]+,", "INDEX\\s*\\([^,]+,\\s*MATCH\\s*\\(\\s*TRUE" };

	public static int Run(string[] args)
	{
		if (args.Length < 1)
		{
			Console.WriteLine("[Help] KimiXlsx recheck <excel_file_path> [--no-recalc] [timeout_seconds]");
			Console.WriteLine("  --no-recalc: Skip LibreOffice recalculation, read cached values directly");
			Console.WriteLine("               Use this to detect errors that only appear in MS Excel");
			return 1;
		}
		string filename = args[0];
		bool skipRecalc = args.Contains("--no-recalc");
		int result;
		int timeout = (from a in args
			where !a.StartsWith("-") && a != filename
			select (!int.TryParse(a, out result)) ? 30 : result).FirstOrDefault(30);
		RecheckResult recheckResult = Recheck(filename, timeout, skipRecalc);
		Console.Write(FormatYamlOutput(recheckResult));
		if (recheckResult.Error == null && recheckResult.ErrorCount <= 0)
		{
			return 0;
		}
		return 1;
	}

	private static string GetMacroDirectory()
	{
		if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
		{
			return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Library/Application Support/LibreOffice/4/user/basic/Standard");
		}
		return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config/libreoffice/4/user/basic/Standard");
	}

	private static string GetMacroFilePath()
	{
		return Path.Combine(GetMacroDirectory(), "Module1.xba");
	}

	private static bool IsMacroConfigured()
	{
		string macroFilePath = GetMacroFilePath();
		if (!File.Exists(macroFilePath))
		{
			return false;
		}
		try
		{
			return File.ReadAllText(macroFilePath).Contains("RecalculateAndSave");
		}
		catch
		{
			return false;
		}
	}

	public static bool SetupLibreOfficeMacro()
	{
		if (IsMacroConfigured())
		{
			return true;
		}
		string macroDirectory = GetMacroDirectory();
		string macroFilePath = GetMacroFilePath();
		if (!Directory.Exists(macroDirectory))
		{
			try
			{
				using Process process = Process.Start(new ProcessStartInfo
				{
					FileName = "soffice",
					Arguments = "--headless --terminate_after_init",
					RedirectStandardOutput = true,
					RedirectStandardError = true,
					UseShellExecute = false
				});
				process?.WaitForExit(10000);
				Directory.CreateDirectory(macroDirectory);
			}
			catch
			{
				return false;
			}
		}
		try
		{
			File.WriteAllText(macroFilePath, "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!DOCTYPE script:module PUBLIC \"-//OpenOffice.org//DTD OfficeDocument 1.0//EN\" \"module.dtd\">\n<script:module xmlns:script=\"http://openoffice.org/2000/script\" script:name=\"Module1\" script:language=\"StarBasic\">\n    Sub RecalculateAndSave()\n      ThisComponent.calculateAll()\n      ThisComponent.store()\n      ThisComponent.close(True)\n    End Sub\n</script:module>");
			return true;
		}
		catch
		{
			return false;
		}
	}

	public static string? RecalculateWithLibreOffice(string absPath, int timeout)
	{
		if (!SetupLibreOfficeMacro())
		{
			return "Failed to setup LibreOffice macro";
		}
		string arguments = "--headless --norestore \"vnd.sun.star.script:Standard.Module1.RecalculateAndSave?language=Basic&location=application\" \"" + absPath + "\"";
		try
		{
			using Process process = Process.Start(new ProcessStartInfo
			{
				FileName = "soffice",
				Arguments = arguments,
				RedirectStandardOutput = true,
				RedirectStandardError = true,
				UseShellExecute = false
			});
			if (process == null)
			{
				return "Failed to start LibreOffice";
			}
			if (!process.WaitForExit(timeout * 1000))
			{
				process.Kill();
				return "LibreOffice timed out";
			}
			if (process.ExitCode != 0 && process.ExitCode != 124)
			{
				string text = process.StandardError.ReadToEnd();
				if (text.Contains("Module1") || !text.Contains("RecalculateAndSave"))
				{
					return "LibreOffice macro not configured properly";
				}
				return text;
			}
			return null;
		}
		catch (Exception ex)
		{
			return ex.Message;
		}
	}

	private static bool IsImplicitArrayFormula(string formula)
	{
		if (string.IsNullOrEmpty(formula))
		{
			return false;
		}
		string text = formula.ToUpperInvariant();
		if (text.Contains("MATCH") && text.Contains("TRUE") && Regex.IsMatch(text, "MATCH\\s*\\(\\s*TRUE\\s*\\(\\s*\\)\\s*,.*[<>=]"))
		{
			return true;
		}
		if (text.Contains("MATCH") && Regex.IsMatch(text, "MATCH\\s*\\([^)]*\\([^)]*[<>=]"))
		{
			return true;
		}
		return false;
	}

	public static (int totalErrors, Dictionary<string, List<string>> errorDetails, int formulaCount, List<string> zeroFormulaCells, List<(string cell, string formula)> implicitArrayFormulas) ScanWithOoxml(string filepath)
	{
		Dictionary<string, List<string>> dictionary = ExcelErrors.ToDictionary((string e) => e, (string e) => new List<string>());
		int num = 0;
		int num2 = 0;
		List<string> list = new List<string>();
		List<(string, string)> list2 = new List<(string, string)>();
		using ZipArchive zipArchive = ZipFile.OpenRead(filepath);
		ZipArchiveEntry entry = zipArchive.GetEntry("xl/workbook.xml");
		ZipArchiveEntry entry2 = zipArchive.GetEntry("xl/_rels/workbook.xml.rels");
		if (entry == null || entry2 == null)
		{
			return (totalErrors: num, errorDetails: dictionary, formulaCount: num2, zeroFormulaCells: list, implicitArrayFormulas: list2);
		}
		XDocument xDocument = XDocument.Load(entry.Open());
		XDocument xDocument2 = XDocument.Load(entry2.Open());
		var list3 = (from s in xDocument.Descendants(MainNs + "sheet")
			select new
			{
				Name = (s.Attribute("name")?.Value ?? ""),
				RId = (s.Attribute(RelNs + "id")?.Value ?? "")
			}).ToList();
		Dictionary<string, string> dictionary2 = (from e in xDocument2.Descendants()
			where e.Name.LocalName == "Relationship"
			select e).ToDictionary((XElement r) => r.Attribute("Id")?.Value ?? "", (XElement r) => r.Attribute("Target")?.Value ?? "");
		foreach (var item in list3)
		{
			if (!dictionary2.TryGetValue(item.RId, out var value))
			{
				continue;
			}
			string entryName = (value.StartsWith("/") ? value.Substring(1) : ("xl/" + value));
			ZipArchiveEntry entry3 = zipArchive.GetEntry(entryName);
			if (entry3 == null)
			{
				continue;
			}
			XElement xElement = XDocument.Load(entry3.Open()).Descendants(MainNs + "sheetData").FirstOrDefault();
			if (xElement == null)
			{
				continue;
			}
			foreach (XElement item2 in xElement.Descendants(MainNs + "row"))
			{
				foreach (XElement item3 in item2.Descendants(MainNs + "c"))
				{
					string text = item3.Attribute("r")?.Value ?? "";
					string text2 = item3.Element(MainNs + "f")?.Value;
					string text3 = item3.Element(MainNs + "v")?.Value;
					if (text2 != null)
					{
						num2++;
						if (text3 != null && double.TryParse(text3, out var result) && result == 0.0)
						{
							list.Add(item.Name + "!" + text);
						}
						if (!(item3.Element(MainNs + "f")?.Attribute("t")?.Value == "array") && IsImplicitArrayFormula(text2))
						{
							list2.Add((item.Name + "!" + text, text2));
						}
					}
					if (text3 == null)
					{
						continue;
					}
					string[] excelErrors = ExcelErrors;
					foreach (string text4 in excelErrors)
					{
						if (text3.Contains(text4))
						{
							dictionary[text4].Add(item.Name + "!" + text);
							num++;
							break;
						}
					}
				}
			}
		}
		return (totalErrors: num, errorDetails: dictionary, formulaCount: num2, zeroFormulaCells: list, implicitArrayFormulas: list2);
	}

	public static RecheckResult Recheck(string filename, int timeout = 30, bool skipRecalc = false)
	{
		if (!File.Exists(filename))
		{
			return new RecheckResult
			{
				Error = "File " + filename + " does not exist"
			};
		}
		string fullPath = Path.GetFullPath(filename);
		if (!skipRecalc)
		{
			string text = RecalculateWithLibreOffice(fullPath, timeout);
			if (text != null)
			{
				return new RecheckResult
				{
					Error = text
				};
			}
		}
		try
		{
			(int totalErrors, Dictionary<string, List<string>> errorDetails, int formulaCount, List<string> zeroFormulaCells, List<(string cell, string formula)> implicitArrayFormulas) tuple = ScanWithOoxml(fullPath);
			int item = tuple.totalErrors;
			Dictionary<string, List<string>> item2 = tuple.errorDetails;
			int item3 = tuple.formulaCount;
			List<string> item4 = tuple.zeroFormulaCells;
			List<(string cell, string formula)> item5 = tuple.implicitArrayFormulas;
			int num = item + item5.Count;
			RecheckResult recheckResult = new RecheckResult
			{
				Result = ((num == 0) ? "pass" : "has_errors"),
				ErrorCount = num,
				FormulaCount = item3,
				ZeroValueCells = new ZeroValueInfo
				{
					Count = item4.Count,
					Cells = item4.Take(20).ToList()
				},
				ImplicitArrayFormulas = new ImplicitArrayFormulaInfo
				{
					Count = item5.Count,
					Cells = (from x in item5.Take(10)
						select x.cell).ToList(),
					Formulas = (from x in item5.Take(10)
						select x.formula).ToList()
				}
			};
			foreach (var (key, list2) in item2)
			{
				if (list2.Count > 0)
				{
					recheckResult.ErrorsByType[key] = new ErrorInfo
					{
						Count = list2.Count,
						Cells = list2.Take(20).ToList()
					};
				}
			}
			return recheckResult;
		}
		catch (Exception ex)
		{
			return new RecheckResult
			{
				Error = ex.Message
			};
		}
	}

	public static string FormatYamlOutput(RecheckResult result)
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
		handler.AppendFormatted(result.Result);
		stringBuilder4.AppendLine(ref handler);
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder5 = stringBuilder2;
		handler = new StringBuilder.AppendInterpolatedStringHandler(15, 1, stringBuilder2);
		handler.AppendLiteral("formula_count: ");
		handler.AppendFormatted(result.FormulaCount);
		stringBuilder5.AppendLine(ref handler);
		stringBuilder2 = stringBuilder;
		StringBuilder stringBuilder6 = stringBuilder2;
		handler = new StringBuilder.AppendInterpolatedStringHandler(13, 1, stringBuilder2);
		handler.AppendLiteral("error_count: ");
		handler.AppendFormatted(result.ErrorCount);
		stringBuilder6.AppendLine(ref handler);
		if (result.ErrorsByType.Count > 0)
		{
			stringBuilder.AppendLine("errors:");
			foreach (KeyValuePair<string, ErrorInfo> item in result.ErrorsByType)
			{
				item.Deconstruct(out var key, out var value);
				string value2 = key;
				ErrorInfo errorInfo = value;
				stringBuilder2 = stringBuilder;
				StringBuilder stringBuilder7 = stringBuilder2;
				handler = new StringBuilder.AppendInterpolatedStringHandler(3, 1, stringBuilder2);
				handler.AppendLiteral("  ");
				handler.AppendFormatted(value2);
				handler.AppendLiteral(":");
				stringBuilder7.AppendLine(ref handler);
				stringBuilder2 = stringBuilder;
				StringBuilder stringBuilder8 = stringBuilder2;
				handler = new StringBuilder.AppendInterpolatedStringHandler(11, 1, stringBuilder2);
				handler.AppendLiteral("    count: ");
				handler.AppendFormatted(errorInfo.Count);
				stringBuilder8.AppendLine(ref handler);
				if (errorInfo.Cells.Count <= 0)
				{
					continue;
				}
				stringBuilder.AppendLine("    cells:");
				foreach (string cell in errorInfo.Cells)
				{
					stringBuilder2 = stringBuilder;
					StringBuilder stringBuilder9 = stringBuilder2;
					handler = new StringBuilder.AppendInterpolatedStringHandler(8, 1, stringBuilder2);
					handler.AppendLiteral("      - ");
					handler.AppendFormatted(cell);
					stringBuilder9.AppendLine(ref handler);
				}
			}
		}
		if (result.ZeroValueCells.Count > 0)
		{
			stringBuilder.AppendLine("zero_value_cells:");
			stringBuilder2 = stringBuilder;
			StringBuilder stringBuilder10 = stringBuilder2;
			handler = new StringBuilder.AppendInterpolatedStringHandler(9, 1, stringBuilder2);
			handler.AppendLiteral("  count: ");
			handler.AppendFormatted(result.ZeroValueCells.Count);
			stringBuilder10.AppendLine(ref handler);
			if (result.ZeroValueCells.Cells.Count > 0)
			{
				stringBuilder.AppendLine("  cells:");
				foreach (string cell2 in result.ZeroValueCells.Cells)
				{
					stringBuilder2 = stringBuilder;
					StringBuilder stringBuilder11 = stringBuilder2;
					handler = new StringBuilder.AppendInterpolatedStringHandler(6, 1, stringBuilder2);
					handler.AppendLiteral("    - ");
					handler.AppendFormatted(cell2);
					stringBuilder11.AppendLine(ref handler);
				}
			}
		}
		if (result.ImplicitArrayFormulas.Count > 0)
		{
			stringBuilder.AppendLine("implicit_array_formulas:");
			stringBuilder2 = stringBuilder;
			StringBuilder stringBuilder12 = stringBuilder2;
			handler = new StringBuilder.AppendInterpolatedStringHandler(9, 1, stringBuilder2);
			handler.AppendLiteral("  count: ");
			handler.AppendFormatted(result.ImplicitArrayFormulas.Count);
			stringBuilder12.AppendLine(ref handler);
			stringBuilder.AppendLine("  note: \"These formulas require CSE (Ctrl+Shift+Enter) in MS Excel but work in LibreOffice. They will show #N/A in Excel!\"");
			if (result.ImplicitArrayFormulas.Cells.Count > 0)
			{
				stringBuilder.AppendLine("  cells:");
				for (int i = 0; i < result.ImplicitArrayFormulas.Cells.Count; i++)
				{
					stringBuilder2 = stringBuilder;
					StringBuilder stringBuilder13 = stringBuilder2;
					handler = new StringBuilder.AppendInterpolatedStringHandler(12, 1, stringBuilder2);
					handler.AppendLiteral("    - cell: ");
					handler.AppendFormatted(result.ImplicitArrayFormulas.Cells[i]);
					stringBuilder13.AppendLine(ref handler);
					stringBuilder2 = stringBuilder;
					StringBuilder stringBuilder14 = stringBuilder2;
					handler = new StringBuilder.AppendInterpolatedStringHandler(17, 1, stringBuilder2);
					handler.AppendLiteral("      formula: \"");
					handler.AppendFormatted(result.ImplicitArrayFormulas.Formulas[i]);
					handler.AppendLiteral("\"");
					stringBuilder14.AppendLine(ref handler);
				}
			}
		}
		return stringBuilder.ToString();
	}
}
