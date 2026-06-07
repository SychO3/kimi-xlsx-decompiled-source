namespace KimiXlsx.Commands;

public class Issue
{
	public string Type { get; set; } = "";

	public string Location { get; set; } = "";

	public string Formula { get; set; } = "";

	public string IssueDescription { get; set; } = "";

	public string? ExpectedPattern { get; set; }

	public string? ActualPattern { get; set; }
}
