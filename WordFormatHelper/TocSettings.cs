namespace WordFormatHelper{

internal struct TocSettings
{
	internal int Levels;

	internal bool UsePageNumber;

	internal int Leader;

	internal int IndentStyle;

	internal int IndentGap;

	internal bool ReplaceCurrentTOC;

	internal bool TryAlignPageNumber;

	public TocSettings()
	{
		Levels = 2;
		UsePageNumber = true;
		Leader = 1;
		IndentStyle = 0;
		IndentGap = 0;
		ReplaceCurrentTOC = true;
		TryAlignPageNumber = false;
	}
}
}