using System.Collections.Generic;

namespace KimiXlsx.Commands;

public class VerificationResult
{
	public string File { get; set; } = "";

	public bool Valid { get; set; } = true;

	public int TotalCharts { get; set; }

	public List<ChartResult> EmptyCharts { get; set; } = new List<ChartResult>();

	public List<ChartResult> ValidCharts { get; set; } = new List<ChartResult>();

	public List<ChartResult> ChartsWithNoData { get; set; } = new List<ChartResult>();

	public List<string> Errors { get; set; } = new List<string>();
}
