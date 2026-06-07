using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class ValidationResult
{
	public int TotalIssues { get; set; }

	public Dictionary<string, int> IssuesByType { get; set; } = new Dictionary<string, int>();

	public List<Issue> Issues { get; set; } = new List<Issue>();

	public string? Error { get; set; }
}
