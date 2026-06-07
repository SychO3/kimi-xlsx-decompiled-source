using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class OpenXmlValidationResult
{
	public string File { get; set; } = "";

	public string Status { get; set; } = "pass";

	public int TotalErrors { get; set; }

	public int IgnoredErrors { get; set; }

	public int CriticalErrors { get; set; }

	public Dictionary<string, int> ErrorsByType { get; set; } = new Dictionary<string, int>();

	public List<ValidationIssue> Errors { get; set; } = new List<ValidationIssue>();

	public List<ValidationIssue> IgnoredIssues { get; set; } = new List<ValidationIssue>();

	public string? FatalError { get; set; }
}
