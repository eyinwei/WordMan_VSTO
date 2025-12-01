using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan
{
    public class TextProcessor
    {
        private Word.ApplicationEvents4_WindowSelectionChangeEventHandler formatPainterSelectionChangeHandler;
        private Microsoft.Office.Tools.Ribbon.RibbonToggleButton currentFormatPainterButton;
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
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                if (sel == null)
                {
                    MessageBox.Show("未检测到可操作的文本或段落。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Word.Range rng;

                if (sel.Range != null &&
                    !string.IsNullOrWhiteSpace(sel.Range.Text) &&
                    sel.Range.Start != sel.Range.End)
                {
                    rng = sel.Range;
                }
                else if (sel.Paragraphs != null && sel.Paragraphs.Count > 0)
                {
                    rng = sel.Paragraphs[1].Range;
                }
                else
                {
                    MessageBox.Show("未检测到可操作的文本或段落。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                rng.Find.Execute(" ", ReplaceWith: "", Replace: Word.WdReplace.wdReplaceAll);
                rng.Find.Execute("　", ReplaceWith: "", Replace: Word.WdReplace.wdReplaceAll);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"去除空格失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region 去除空行功能
        public void RemoveEmptyLines()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                if (sel == null || sel.Range == null)
                {
                    MessageBox.Show("未检测到可操作的文本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Word.Range rng = sel.Range;

                // 从后往前遍历选区内的所有段落
                for (int i = rng.Paragraphs.Count; i >= 1; i--)
                {
                    Word.Paragraph para = rng.Paragraphs[i];
                    if (para == null || para.Range == null)
                        continue;

                    // 去除回车、换行、空格、全角空格、Tab等
                    string text = para.Range.Text.Trim('\r', '\n', ' ', '\t', '　');
                    if (string.IsNullOrEmpty(text))
                    {
                        para.Range.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"去除空行失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    MessageBox.Show("未检测到可操作的文本或段落。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    {"　", " "}, {"＆", "&"}, {"＃", "#"},
                    {"％", "%"}, {"＊", "*"}, {"＋", "+"}, {"－", "-"}, {"＝", "="},
                    {"＠", "@"}, {"＄", "$"}, {"＾", "^"}, {"＿", "_"}, {"｀", "`"},
                    {"｜", "|"}, {"＼", "\\"}, {"～", "~"}
                };

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

            Regex regex = new Regex(pattern);
            string text = range.Text;

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
                    insertRange.InsertAfter(" ");
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
                paraFormat.CharacterUnitFirstLineIndent = 0f;
                paraFormat.FirstLineIndent = 0f;
                paraFormat.CharacterUnitLeftIndent = 0f;
                paraFormat.LeftIndent = 0f;
                paraFormat.CharacterUnitRightIndent = 0f;
                paraFormat.RightIndent = 0f;
            }
        }
        #endregion

        #region 字体替换功能
        /// <summary>
        /// 替换指定字体
        /// </summary>
        /// <param name="originalFont">原字体名称</param>
        /// <param name="newFont">新字体名称</param>
        public void ReplaceFont(string originalFont, string newFont)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            Word.Range rng;

            if (sel != null && sel.Range != null && sel.Range.Start != sel.Range.End)
            {
                rng = sel.Range.Duplicate;
            }
            else
            {
                rng = app.ActiveDocument.Range();
            }

            rng.Find.ClearFormatting();
            rng.Find.Replacement.ClearFormatting();
            rng.Find.Font.Name = originalFont;
            rng.Find.Replacement.Font.Name = newFont;
            rng.Find.Text = "";
            rng.Find.Replacement.Text = "";
            rng.Find.Forward = true;
            rng.Find.Wrap = Word.WdFindWrap.wdFindStop;
            rng.Find.Format = true;
            rng.Find.MatchCase = false;
            rng.Find.MatchWholeWord = false;
            rng.Find.MatchWildcards = false;
            rng.Find.MatchSoundsLike = false;
            rng.Find.MatchAllWordForms = false;
            rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }

        /// <summary>
        /// 将仿宋GB2312字体替换为仿宋字体
        /// </summary>
        public void ReplaceFangSongGB2312ToFangSong()
        {
            ReplaceFont("仿宋_GB2312", "仿宋");
        }

        /// <summary>
        /// 将楷体GB2312字体替换为楷体字体
        /// </summary>
        public void ReplaceKaiTiGB2312ToKaiTi()
        {
            ReplaceFont("楷体_GB2312", "楷体");
        }

        /// <summary>
        /// 将方正小标宋简体字体替换为黑体字体
        /// </summary>
        public void ReplaceFZXBSToHeiTi()
        {
            ReplaceFont("方正小标宋简体", "黑体");
        }

        /// <summary>
        /// 将数字、英文大小写字母等替换为Times New Roman字体
        /// </summary>
        public void ReplaceAllToTimesNewRoman()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                Word.Range rng;

                if (sel != null && sel.Range != null && sel.Range.Start != sel.Range.End)
                {
                    rng = sel.Range.Duplicate;
                }
                else
                {
                    rng = app.ActiveDocument.Range();
                }

                if (rng == null)
                {
                    return;
                }

                app.ScreenUpdating = false;

                rng.Find.ClearFormatting();
                rng.Find.Replacement.ClearFormatting();
                rng.Find.Replacement.Font.Name = "Times New Roman";
                rng.Find.Text = "[0-9A-Za-z]";
                rng.Find.Replacement.Text = "^&";
                rng.Find.Forward = true;
                rng.Find.Wrap = Word.WdFindWrap.wdFindStop;
                rng.Find.Format = true;
                rng.Find.MatchCase = false;
                rng.Find.MatchWholeWord = false;
                rng.Find.MatchWildcards = true;
                rng.Find.MatchSoundsLike = false;
                rng.Find.MatchAllWordForms = false;
                rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }

        #endregion

        #region Word 内置功能
        /// <summary>
        /// 清除格式
        /// </summary>
        public void ClearFormatting()
        {
            try
            {
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("ClearFormatting");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行清除格式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 格式刷点击处理
        /// </summary>
        /// <param name="toggleButton">格式刷切换按钮</param>
        public void FormatPainter_Click(Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (toggleButton.Checked)
                {
                    // 如果之前有事件处理器，先移除
                    if (formatPainterSelectionChangeHandler != null)
                    {
                        app.WindowSelectionChange -= formatPainterSelectionChangeHandler;
                    }
                    
                    // 激活格式刷
                    app.CommandBars.ExecuteMso("FormatPainter");
                    
                    // 保存当前按钮引用
                    currentFormatPainterButton = toggleButton;
                    
                    // 创建并保存事件处理器
                    formatPainterSelectionChangeHandler = new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(FormatPainter_SelectionChange);
                    
                    // 监听选择变化，当格式应用后自动取消选中
                    app.WindowSelectionChange += formatPainterSelectionChangeHandler;
                }
                else
                {
                    // 取消格式刷模式
                    if (formatPainterSelectionChangeHandler != null)
                    {
                        app.WindowSelectionChange -= formatPainterSelectionChangeHandler;
                        formatPainterSelectionChangeHandler = null;
                        currentFormatPainterButton = null;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行格式刷失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 格式刷选择变化处理
        /// </summary>
        /// <param name="sel">当前选择</param>
        private void FormatPainter_SelectionChange(Word.Selection sel)
        {
            try
            {
                // 当选择变化时，如果格式刷已经应用，自动取消选中
                if (currentFormatPainterButton != null && currentFormatPainterButton.Checked)
                {
                    var app = Globals.ThisAddIn.Application;
                    // 使用 Timer 延迟检查，避免立即触发
                    System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
                    timer.Interval = 200;
                    timer.Tick += (s, args) =>
                    {
                        timer.Stop();
                        timer.Dispose();
                        if (currentFormatPainterButton != null && currentFormatPainterButton.Checked)
                        {
                            currentFormatPainterButton.Checked = false;
                            if (formatPainterSelectionChangeHandler != null)
                            {
                                app.WindowSelectionChange -= formatPainterSelectionChangeHandler;
                                formatPainterSelectionChangeHandler = null;
                                currentFormatPainterButton = null;
                            }
                        }
                    };
                    timer.Start();
                }
            }
            catch { }
        }

        /// <summary>
        /// 只留文本（粘贴为纯文本或清除格式）
        /// </summary>
        public void PasteTextOnly()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                
                // 方法1：优先尝试使用 Word 内置的"只粘贴文本"命令
                try
                {
                    app.CommandBars.ExecuteMso("PasteTextOnly");
                    return;
                }
                catch
                {
                    // 如果命令不存在，使用 PasteSpecial 方法
                }
                
                // 方法2：使用 Word 的 PasteSpecial 方法粘贴为纯文本
                try
                {
                    sel.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteText);
                    return;
                }
                catch (Exception pasteEx)
                {
                    // 如果 PasteSpecial 也失败，提示用户
                    if (sel != null && sel.Type != Word.WdSelectionType.wdSelectionIP)
                    {
                        // 如果已选中文本，清除格式
                        sel.ClearFormatting();
                    }
                    else
                    {
                        MessageBox.Show($"无法执行只粘贴文本操作：{pasteEx.Message}\n\n请确保剪贴板中有文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行只留文本失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

    }
}