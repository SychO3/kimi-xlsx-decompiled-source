using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class ChartResult
{
	public string File { get; set; } = "";

	public string Name { get; set; } = "";

	public bool HasData { get; set; }

	public string? ChartType { get; set; }

	public int SeriesCount { get; set; }

	public int DataPoints { get; set; }

	public int NonZeroDataPoints { get; set; }

	public int CategoryCount { get; set; }

	public bool HasCategories { get; set; }

	public bool HasValues { get; set; }

	public bool HasMeaningfulData { get; set; }

	public List<string> Issues { get; set; } = new List<string>();
}
