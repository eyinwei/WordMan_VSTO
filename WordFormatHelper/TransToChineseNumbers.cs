using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace WordFormatHelper{

public class TransToChineseNumbers
{
	private readonly List<string> Chs_Numbers = new List<string>(10) { "〇", "一", "二", "三", "四", "五", "六", "七", "八", "九" };

	private readonly List<string> Chs_NumberSys = new List<string>(4) { "", "十", "百", "千" };

	private readonly List<string> Chs_NumberSys1 = new List<string>(5) { "", "万", "亿", "兆", "京" };

	private readonly List<string> Cht_Numbers = new List<string>(10) { "零", "壹", "貳", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };

	private readonly List<string> Cht_NumberSys = new List<string>(4) { "", "拾", "佰", "仟" };

	private readonly List<string> Cht_NumberSys1 = new List<string>(5) { "", "萬", "億", "兆", "京" };

	public string ToChineseNumber(string ArabicNumbers, [Optional] bool ChineseTraditional)
	{
		try
		{
			ArabicNumbers = ArabicNumbers.Trim(' ', '\n', '\r');
			ArabicNumbers = ArabicNumbers.Replace(",", "");
			if (Math.Abs(Convert.ToDouble(ArabicNumbers)) >= 1E+21)
			{
				throw new Exception("转换数值超出范围。");
			}
		}
		catch (Exception ex)
		{
			return ex.Message;
		}
		string text = "";
		string text2 = "";
		string text3 = "";
		if (ArabicNumbers.StartsWith("-"))
		{
			ArabicNumbers = ArabicNumbers.TrimStart('-');
			text2 = (ChineseTraditional ? "負" : "负");
		}
		string text4;
		if (ArabicNumbers.Contains("."))
		{
			text4 = ArabicNumbers.Split('.')[0];
			text = ArabicNumbers.Split('.')[1];
			text3 = (ChineseTraditional ? "點" : "点");
		}
		else
		{
			text4 = ArabicNumbers;
		}
		List<string> list;
		List<string> list2;
		List<string> list3;
		if (ChineseTraditional)
		{
			list = Cht_Numbers;
			list2 = Cht_NumberSys;
			list3 = Cht_NumberSys1;
		}
		else
		{
			list = Chs_Numbers;
			list2 = Chs_NumberSys;
			list3 = Chs_NumberSys1;
		}
		string text5 = "";
		for (int i = 0; i < 10; i++)
		{
			text4 = text4.Replace(i.ToString(), list[i]);
			text = text.Replace(i.ToString(), list[i]);
		}
		int num = 0;
		while (num < text4.Length)
		{
			string text6 = "";
			for (int j = 0; j < 4; j++)
			{
				if (num == text4.Length)
				{
					break;
				}
				string text7 = text4.Substring(text4.Length - 1 - num, 1);
				text6 = ((!(text7 != "〇")) ? (text7 + text6) : (text7 + list2[j] + text6));
				num++;
			}
			text6 = Regex.Replace(text6, "〇{2,}", "〇");
			text6 = text6.TrimEnd('〇');
			if (text6 != "")
			{
				text5 = text6 + list3[(num - 1) / 4] + text5;
			}
		}
		if (text5.StartsWith(list[1] + list2[1]))
		{
			text5 = text5.TrimStart(list[1].ToCharArray());
		}
		if (text5.StartsWith(list[0]))
		{
			text5 = text5.TrimStart(list[0].ToCharArray());
		}
		return text2 + text5 + text3 + text;
	}
}
}