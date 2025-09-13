namespace WordFormatHelper{

internal struct HeaderFooterFontInfo
{
	internal string HeaderFontName;

	internal float HeaderFontSize;

	internal bool HeaderFontBold;

	internal bool HeaderFontItalic;

	internal string FooterFontName;

	internal float FooterFontSize;

	internal bool FooterFontBold;

	internal bool FooterFontItalic;

	public HeaderFooterFontInfo()
	{
		HeaderFontSize = 0f;
		HeaderFontBold = false;
		HeaderFontItalic = false;
		FooterFontSize = 0f;
		FooterFontBold = false;
		FooterFontItalic = false;
		HeaderFontName = string.Empty;
		FooterFontName = string.Empty;
	}
}
}