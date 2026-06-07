using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class ImplicitArrayFormulaInfo
{
	public int Count { get; set; }

	public List<string> Cells { get; set; } = new List<string>();

	public List<string> Formulas { get; set; } = new List<string>();
}
