using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan
{
    public class TextProcessor
    {
        #region 去除断行功能
        public void RemoveLineBreaks()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            if (sel == null || sel.Range == null || string.IsNullOrEmpty(sel.Range.Text))
                return;

            Word.Range rng = sel.Range.Duplicate;
            string text = rng.Text;

            // 判断末尾是否有回车
            bool endsWithReturn = text.EndsWith("\r");

            // 如果结尾有回车，先排除最后一个回车后再处理
            int processLength = endsWithReturn ? text.Length - 1 : text.Length;
            Word.Range processRange = rng.Duplicate;
            processRange.End = processRange.Start + processLength;

            // 替换所有回车
            processRange.Find.ClearFormatting();
            processRange.Find.Replacement.ClearFormatting();
            processRange.Find.Text = "^013"; // 回车
            processRange.Find.Replacement.Text = "";
            processRange.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);

            // 替换所有软回车
            processRange.Find.Text = "^11"; // 手动换行(软回车)
            processRange.Find.Replacement.Text = "";
            processRange.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);

            // 这样可一键撤销，且格式不会丢失
        }
        #endregion

        #region 去除空格功能
        public void RemoveSpaces()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            Word.Range rng;

            if (sel != null && sel.Range != null &&
                !string.IsNullOrWhiteSpace(sel.Range.Text) &&
                sel.Range.Start != sel.Range.End)
            {
                rng = sel.Range;
            }
            else if (sel != null && sel.Paragraphs != null && sel.Paragraphs.Count > 0)
            {
                rng = sel.Paragraphs[1].Range;
            }
            else
            {
                MessageBox.Show("未检测到可操作的文本或段落。");
                return;
            }
            rng.Find.Execute(" ", ReplaceWith: "", Replace: Word.WdReplace.wdReplaceAll);
            rng.Find.Execute("　", ReplaceWith: "", Replace: Word.WdReplace.wdReplaceAll);
        }
        #endregion

        #region 去除空行功能
        public void RemoveEmptyLines()
        {
            // 获取Word应用程序对象
            var app = Globals.ThisAddIn.Application;

            // 获取当前选区
            Word.Range rng = app.Selection.Range;

            // 从后往前遍历选区内的所有段落
            for (int i = rng.Paragraphs.Count; i >= 1; i--)
            {
                Word.Paragraph para = rng.Paragraphs[i];
                // 去除回车、换行、空格、全角空格、Tab等
                string text = para.Range.Text.Trim('\r', '\n', ' ', '\t', '　');
                if (string.IsNullOrEmpty(text))
                {
                    para.Range.Delete();
                }
            }
        }
        #endregion

        #region 符号转换功能
        public void ConvertPunctuation(bool englishToChinese)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var rng = app.Selection.Range.Start != app.Selection.Range.End
                    ? app.Selection.Range
                    : app.Selection.Paragraphs[1].Range;

                if (string.IsNullOrWhiteSpace(rng.Text))
                {
                    MessageBox.Show("未检测到可操作的文本或段落。");
                    return;
                }

                app.ScreenUpdating = false;

                // 标点符号映射
                var dict = englishToChinese ? new Dictionary<string, string>
                {
                    {";", "；"}, {":", "："}, {",", "，"}, {".", "。"}, {"?", "？"}, {"!", "！"},
                    {"(", "（"}, {")", "）"}, {"[", "【"}, {"]", "】"}, {"<", "《"}, {">", "》"}
                } : new Dictionary<string, string>
                {
                    {"；", ";"}, {"：", ":"}, {"，", ","}, {"。", "."}, {"？", "?"}, {"！", "!"},
                    {"（", "("}, {"）", ")"}, {"【", "["}, {"】", "]"}, {"《", "<"}, {"》", ">"},
                    {"　", " "}, {"＂", "\""}, {"＇", "'"}, {"＆", "&"}, {"＃", "#"},
                    {"％", "%"}, {"＊", "*"}, {"＋", "+"}, {"－", "-"}, {"＝", "="},
                    {"＠", "@"}, {"＄", "$"}, {"＾", "^"}, {"＿", "_"}, {"｀", "`"},
                    {"｜", "|"}, {"＼", "\\"}, {"～", "~"}
                };

                // 使用Word查找替换保持格式
                // 使用范围副本避免原始范围被修改
                Word.Range workingRange = rng.Duplicate;
                foreach (var pair in dict)
                {
                    try
                    {
                        workingRange.Find.ClearFormatting();
                        workingRange.Find.Replacement.ClearFormatting();
                        workingRange.Find.Text = pair.Key;
                        workingRange.Find.Replacement.Text = pair.Value;
                        workingRange.Find.MatchWildcards = false;
                        workingRange.Find.MatchCase = false;
                        workingRange.Find.MatchWholeWord = false;
                        workingRange.Find.Forward = true;
                        workingRange.Find.Wrap = Word.WdFindWrap.wdFindStop;
                        workingRange.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                    }
                    catch
                    {
                        // 忽略单个符号替换失败，继续处理其他符号
                        continue;
                    }
                }

                // 英标转中标时处理成对引号
                if (englishToChinese)
                {
                    try
                    {
                        // 保存原始范围
                        int startPos = rng.Start;
                        int endPos = rng.End;

                        // 先收集所有引号位置，避免在循环中修改文档导致位置变化
                        List<int> doubleQuotePositions = new List<int>();
                        List<int> singleQuotePositions = new List<int>();

                        // 收集双引号位置（使用字符串搜索，避免Find的复杂性）
                        string rangeText = app.ActiveDocument.Range(startPos, endPos).Text;
                        int currentPos = startPos;
                        int textIndex = 0;
                        while (textIndex < rangeText.Length)
                        {
                            if (rangeText[textIndex] == '"')
                            {
                                doubleQuotePositions.Add(currentPos + textIndex);
                            }
                            textIndex++;
                        }

                        // 收集单引号位置
                        textIndex = 0;
                        while (textIndex < rangeText.Length)
                        {
                            if (rangeText[textIndex] == '\'')
                            {
                                singleQuotePositions.Add(currentPos + textIndex);
                            }
                            textIndex++;
                        }

                        // 倒序替换双引号（避免位置变化影响后续替换）
                        bool isLeft = true;
                        for (int i = doubleQuotePositions.Count - 1; i >= 0; i--)
                        {
                            try
                            {
                                var hitRange = app.ActiveDocument.Range(doubleQuotePositions[i], doubleQuotePositions[i] + 1);
                                string currentText = hitRange.Text;
                                if (currentText == "\"")
                                {
                                    hitRange.Text = isLeft ? "\u201c" : "\u201d";
                                    isLeft = !isLeft;
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }

                        // 倒序替换单引号
                        isLeft = true;
                        for (int i = singleQuotePositions.Count - 1; i >= 0; i--)
                        {
                            try
                            {
                                var hitRange = app.ActiveDocument.Range(singleQuotePositions[i], singleQuotePositions[i] + 1);
                                string currentText = hitRange.Text;
                                if (currentText == "'")
                                {
                                    hitRange.Text = isLeft ? "\u2018" : "\u2019";
                                    isLeft = !isLeft;
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                    catch
                    {
                        // 忽略引号替换失败
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"转换失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }

        public void ConvertEnglishToChinesePunctuation()
        {
            ConvertPunctuation(true);
        }

        public void ConvertChineseToEnglishPunctuation()
        {
            ConvertPunctuation(false);
        }
        #endregion

        #region 自动加空格功能
        public void AutoAddSpaces()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app.Selection;

            // 需要查找的正则：数字后紧跟单位（英文字母、μ、Ω、°、℃等），且二者之间无空格
            string pattern = @"(\d+(?:\.\d+)?)([a-zA-ZμΩℓ‰°℃℉Å])";

            // 只处理选区或当前段落
            Word.Range range = selection.Type == Word.WdSelectionType.wdSelectionNormal
                ? selection.Range
                : selection.Paragraphs[1].Range;

            // 由于Word原生Find不支持复杂正则，所以采用文本查找+偏移定位方式
            Regex regex = new Regex(pattern);
            string text = range.Text;

            // 记录需要插入空格的相对位置（倒序处理，防止位置错乱）
            var matches = regex.Matches(text);
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                var match = matches[i];
                int insertPos = range.Start + match.Index + match.Groups[1].Length;
                // 检查当前位置是否已有空格
                if (text.Length > match.Index + match.Groups[1].Length
                    && text[match.Index + match.Groups[1].Length] != ' ')
                {
                    Word.Range insertRange = range.Duplicate;
                    insertRange.Start = insertRange.End = insertPos;
                    insertRange.InsertAfter(" "); // 在数字和单位中间插入空格，样式不变
                }
            }
        }
        #endregion

        #region 缩进功能
        public void IndentTwoCharacters()
        {
            var selection = Globals.ThisAddIn.Application.Selection;
            if (selection != null)
            {
                var paraFormat = selection.ParagraphFormat;
                paraFormat.CharacterUnitFirstLineIndent = 2f;
            }
        }

        public void RemoveIndent()
        {
            var selection = Globals.ThisAddIn.Application.Selection;
            if (selection != null)
            {
                var paraFormat = selection.ParagraphFormat;

                // 先清除首行缩进（字符和磅）
                paraFormat.CharacterUnitFirstLineIndent = 0f;
                paraFormat.FirstLineIndent = 0f;

                // 再清除段落左缩进（字符和磅）
                paraFormat.CharacterUnitLeftIndent = 0f;
                paraFormat.LeftIndent = 0f;

                // 可选：右缩进也清零（防止有些文档右缩进影响排版）
                paraFormat.CharacterUnitRightIndent = 0f;
                paraFormat.RightIndent = 0f;
            }
        }
        #endregion
    }
}