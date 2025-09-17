// WordAssistant, Version=1.3.5.0, Culture=neutral, PublicKeyToken=null
// WordFormatHelper.Globals
using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using Microsoft.Office.Tools.Word;
using WordFormatHelper;

[DebuggerNonUserCode]
[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "17.0.0.0")]
internal sealed class Globals
{
	private static ThisAddIn _ThisAddIn;

	private static ApplicationFactory _factory;

	private static ThisRibbonCollection _ThisRibbonCollection;

	internal static ThisAddIn ThisAddIn
	{
		get
		{
			return _ThisAddIn;
		}
		set
		{
			if (_ThisAddIn == null)
			{
				_ThisAddIn = value;
				return;
			}
			throw new NotSupportedException();
		}
	}

	internal static ApplicationFactory Factory
	{
		get
		{
			return _factory;
		}
		set
		{
			if (_factory == null)
			{
				_factory = value;
				return;
			}
			throw new NotSupportedException();
		}
	}

	internal static ThisRibbonCollection Ribbons
	{
		get
		{
			if (_ThisRibbonCollection == null)
			{
				_ThisRibbonCollection = new ThisRibbonCollection(_factory.GetRibbonFactory());
			}
			return _ThisRibbonCollection;
		}
	}

	private Globals()
	{
	}
}
