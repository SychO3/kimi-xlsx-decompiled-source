using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class InspectionResult
{
	public string Status { get; set; } = "success";

	public string File { get; set; } = "";

	public List<SheetInfo> Sheets { get; set; } = new List<SheetInfo>();

	public string? Message { get; set; }
}
