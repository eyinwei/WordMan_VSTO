using System;
using Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式信息类
    /// 用于存储和序列化样式设置
    /// </summary>
    [Serializable]
    public class StyleInfo
    {
        public string StyleName { get; set; }
        public string ChnFontName { get; set; }
        public string EngFontName { get; set; }
        public string FontSize { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string FontColor { get; set; }
        public WdParagraphAlignment Alignment { get; set; }
        public float SpaceBefore { get; set; }
        public float SpaceAfter { get; set; }
        public float LineSpacing { get; set; }
        public float FirstLineIndent { get; set; }
        public bool PageBreakBefore { get; set; }
        public bool IsBuiltIn { get; set; }

        public StyleInfo()
        {
            StyleName = "";
            ChnFontName = "仿宋";
            EngFontName = "仿宋";
            FontSize = "16";
            Bold = false;
            Italic = false;
            Underline = false;
            FontColor = "#FF000000";
            Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            SpaceBefore = 0f;
            SpaceAfter = 0f;
            LineSpacing = 28f;
            FirstLineIndent = 0f;
            PageBreakBefore = false;
            IsBuiltIn = false;
        }

        public StyleInfo(string styleName) : this()
        {
            StyleName = styleName;
        }

        /// <summary>
        /// 从Hashtable创建StyleInfo
        /// </summary>
        /// <param name="styleName">样式名称</param>
        /// <param name="settings">样式设置</param>
        /// <returns>StyleInfo对象</returns>
        public static StyleInfo FromHashtable(string styleName, System.Collections.Hashtable settings)
        {
            var styleInfo = new StyleInfo(styleName);
            
            if (settings != null)
            {
                styleInfo.ChnFontName = settings["cnFont"]?.ToString() ?? "仿宋";
                styleInfo.EngFontName = settings["enFont"]?.ToString() ?? "仿宋";
                styleInfo.FontSize = settings["fontSize"]?.ToString() ?? "16";
                styleInfo.Bold = Convert.ToBoolean(settings["isBold"]);
                styleInfo.Alignment = (WdParagraphAlignment)(settings["alignment"] ?? WdParagraphAlignment.wdAlignParagraphLeft);
                styleInfo.SpaceBefore = Convert.ToSingle(settings["spaceBefore"] ?? 0f);
                styleInfo.SpaceAfter = Convert.ToSingle(settings["spaceAfter"] ?? 0f);
                styleInfo.LineSpacing = Convert.ToSingle(settings["lineSpacing"] ?? 28f);
                
                if (settings.ContainsKey("firstLineIndent"))
                {
                    styleInfo.FirstLineIndent = Convert.ToSingle(settings["firstLineIndent"]);
                }
            }
            
            return styleInfo;
        }

        /// <summary>
        /// 转换为Hashtable
        /// </summary>
        /// <returns>Hashtable对象</returns>
        public System.Collections.Hashtable ToHashtable()
        {
            return new System.Collections.Hashtable
            {
                {"cnFont", ChnFontName},
                {"enFont", EngFontName},
                {"fontSize", FontSize},
                {"isBold", Bold},
                {"alignment", Alignment},
                {"spaceBefore", SpaceBefore},
                {"spaceAfter", SpaceAfter},
                {"lineSpacing", LineSpacing},
                {"firstLineIndent", FirstLineIndent}
            };
        }
    }
}
