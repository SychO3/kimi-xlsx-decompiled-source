using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class SheetInfo
{
	public string Name { get; set; } = "";

	public List<TableInfo> Tables { get; set; } = new List<TableInfo>();

	public int TotalRows { get; set; }

	public int TotalCols { get; set; }

	public string? DataRange { get; set; }

	public string? Error { get; set; }
}
