using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class RecheckResult
{
	public string Result { get; set; } = "pass";

	public int ErrorCount { get; set; }

	public int FormulaCount { get; set; }

	public Dictionary<string, ErrorInfo> ErrorsByType { get; set; } = new Dictionary<string, ErrorInfo>();

	public ZeroValueInfo ZeroValueCells { get; set; } = new ZeroValueInfo();

	public ImplicitArrayFormulaInfo ImplicitArrayFormulas { get; set; } = new ImplicitArrayFormulaInfo();

	public string? Error { get; set; }
}
