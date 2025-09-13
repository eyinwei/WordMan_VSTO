using System.CodeDom.Compiler;
using System.Diagnostics;
using Microsoft.Office.Tools.Ribbon;

namespace WordFormatHelper{

[DebuggerNonUserCode]
[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
internal sealed class ThisRibbonCollection : RibbonCollectionBase
{
	internal WordFormatHelperRibbon WordFormatHelperRibbon => GetRibbon<WordFormatHelperRibbon>();

	internal ThisRibbonCollection(RibbonFactory factory)
		: base(factory)
	{
	}
}
}