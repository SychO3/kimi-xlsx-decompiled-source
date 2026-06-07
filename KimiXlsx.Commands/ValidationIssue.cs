namespace KimiXlsx.Commands;

public class ValidationIssue
{
	public string ErrorType { get; set; } = "";

	public string Description { get; set; } = "";

	public string? Path { get; set; }

	public string? Part { get; set; }

	public bool Ignored { get; set; }

	public string? IgnoreReason { get; set; }
}
