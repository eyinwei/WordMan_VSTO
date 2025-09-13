namespace WordFormatHelper{

internal struct HeaderFooterTextInfo
{
	internal string[] PrimaryHeaderText;

	internal string[] FirstHeaderText;

	internal string[] EvenHeaderText;

	internal string[] PrimaryFooterText;

	internal string[] FirstFooterText;

	internal string[] EvenFooterText;

	internal bool FirstPageDiffrent;

	internal bool OddEvenPageDiffrent;

	internal int HeaderLineType;

	internal int FooterLineType;

	internal string[] LogoPath;

	internal float LogoHeight;

	internal int ApplyModel;

	internal bool PageNumberStartAtSection;

	internal int StartNumber;

	internal bool SameHeaderFooterHeight;

	public HeaderFooterTextInfo()
	{
		PrimaryHeaderText = new string[3] { "", "", "" };
		FirstHeaderText = new string[3] { "", "", "" };
		EvenHeaderText = new string[3] { "", "", "" };
		PrimaryFooterText = new string[3] { "", "", "" };
		FirstFooterText = new string[3] { "", "", "" };
		EvenFooterText = new string[3] { "", "", "" };
		FirstPageDiffrent = false;
		OddEvenPageDiffrent = false;
		HeaderLineType = 5;
		FooterLineType = 5;
		LogoPath = new string[3] { "", "", "" };
		LogoHeight = 1f;
		ApplyModel = 0;
		PageNumberStartAtSection = false;
		StartNumber = 1;
		SameHeaderFooterHeight = false;
	}
}
}