using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class TableInfo
{
	public string Range { get; set; } = "";

	public int StartRow { get; set; }

	public int EndRow { get; set; }

	public string StartCol { get; set; } = "";

	public string EndCol { get; set; } = "";

	public int RowCount { get; set; }

	public int ColCount { get; set; }

	public int? HeaderRow { get; set; }

	public List<string> Headers { get; set; } = new List<string>();

	public int DataStartRow { get; set; }

	public int DataEndRow { get; set; }

	public int? TotalRow { get; set; }
}
