using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace WordFormatHelper{

[Serializable]
public class FixFormatPainterSetting
{
	public struct FixFormat
	{
		public int Id { get; set; }

		public string StyleName { get; set; }

		public string Discription { get; set; }

		public string EngFontName { get; set; }

		public string ChnFontName { get; set; }

		public float FontSize { get; set; }

		public bool Bold { get; set; }

		public bool Italic { get; set; }

		public bool Underline { get; set; }

		public bool UseColor { get; set; }

		public int TextColor { get; set; }

		public bool Shading { get; set; }

		public int ShadingColor { get; set; }
	}

	public int CurrentID;

	public List<FixFormat> StoredFormat = new List<FixFormat>();

	public FixFormatPainterSetting()
	{
		StoredFormat = new List<FixFormat>();
	}

	public static FixFormatPainterSetting FromXmlFile(string filePath)
	{
		if (!File.Exists(filePath))
		{
			throw new FileNotFoundException("未能读取配置文件！", filePath);
		}
		XmlSerializer xmlSerializer = new XmlSerializer(typeof(FixFormatPainterSetting));
		using StreamReader textReader = new StreamReader(filePath);
		return (FixFormatPainterSetting)xmlSerializer.Deserialize(textReader);
	}

	public void ToXmlFile(string filePath)
	{
		if (StoredFormat.Count == 0)
		{
			throw new Exception("当前实例不包含任何数据！");
		}
		XmlSerializer xmlSerializer = new XmlSerializer(typeof(FixFormatPainterSetting));
		using StreamWriter textWriter = new StreamWriter(filePath);
		xmlSerializer.Serialize(textWriter, this);
	}
}
}