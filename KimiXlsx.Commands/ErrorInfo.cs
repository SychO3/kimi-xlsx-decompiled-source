using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class ErrorInfo
{
	public int Count { get; set; }

	public List<string> Cells { get; set; } = new List<string>();
}
