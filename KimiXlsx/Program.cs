using System;
using System.Linq;
using KimiXlsx.Commands;

namespace KimiXlsx;

internal class Program
{
	private static void PrintHelp()
	{
		Console.WriteLine("\nKimiXlsx - Excel Manipulation Tool (.NET 8.0)\n=============================================\n\nUsage: dotnet run --project KimiXlsx -- <command> [arguments]\n\nCommands:\n  validate <file.xlsx> [--json] [--strict] [--delete-invalid]\n      ⚠\ufe0f  MANDATORY: OpenXML structure validation.\n      Files that fail this check CANNOT be opened by MS Excel.\n      Run this BEFORE delivering any Excel file.\n\n  pivot <input.xlsx> <output.xlsx> --source --location --values [options]\n      Create PivotTable using pure OpenXML SDK.\n      Required: --source \"Sheet!A1:Z100\" --location \"Sheet!A3\" --values \"Field:sum\"\n      Optional: --rows, --cols, --filters, --name\n\n  reference-check <file.xlsx>\n      Check formula references for errors and consistency issues.\n\n  recheck <file.xlsx> [timeout_seconds]\n      Recalculate formulas via LibreOffice and scan for errors.\n      Requires LibreOffice installed with 'soffice' in PATH.\n\n  inspect <file.xlsx> [--pretty/-p] [--output/-o <file.json>]\n      Analyze Excel file structure (sheets, tables, headers).\n\n  chart-verify <file.xlsx> [--verbose/-v] [--json]\n      Verify that all charts have actual data content.\n\n  help\n      Show this help message.\n\nExamples:\n  dotnet run --project KimiXlsx -- reference-check data.xlsx\n  dotnet run --project KimiXlsx -- inspect data.xlsx --pretty\n");
	}

	private static int ShowHelp()
	{
		PrintHelp();
		return 0;
	}

	private static int UnknownCommand(string command)
	{
		Console.WriteLine("Unknown command: " + command);
		PrintHelp();
		return 1;
	}

	private static int Main(string[] args)
	{
		if (args.Length == 0)
		{
			PrintHelp();
			return 1;
		}
		string text = args[0].ToLower();
		string[] args2 = args.Skip(1).ToArray();
		switch (text)
		{
		case "validate":
			return ValidateCommand.Run(args2);
		case "pivot":
			return PivotCommand.Run(args2);
		case "refcheck":
		case "reference-check":
			return ReferenceCheckCommand.Run(args2);
		case "recheck":
			return RecheckCommand.Run(args2);
		case "inspect":
			return InspectCommand.Run(args2);
		case "chart-verify":
		case "chartverify":
			return ChartVerifyCommand.Run(args2);
		case "help":
		case "--help":
		case "-h":
			return ShowHelp();
		default:
			return UnknownCommand(text);
		}
	}
}
