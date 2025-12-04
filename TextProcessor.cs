using System;
using System.Collections.Generic;
using System.Linq;
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

                Word.Range rng = null;

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

                if (rng == null)
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
                var sel = app.Selection;
                Word.Range rng = null;

                if (sel != null && sel.Range != null && sel.Range.Start != sel.Range.End)
                {
                    rng = sel.Range;
                }
                else if (sel != null && sel.Paragraphs != null && sel.Paragraphs.Count > 0)
                {
                    rng = sel.Paragraphs[1].Range;
                }

                if (rng == null || string.IsNullOrWhiteSpace(rng.Text))
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
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.Selection;

                if (selection == null)
                {
                    MessageBox.Show("未检测到可操作的文本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Word.Range range = null;
                
                // 确定处理范围：优先使用选区，否则使用当前段落
                if (selection.Type == Word.WdSelectionType.wdSelectionNormal
                    && selection.Range != null
                    && selection.Range.Start != selection.Range.End)
                {
                    range = selection.Range;
                }
                else if (selection.Paragraphs != null && selection.Paragraphs.Count > 0)
                {
                    range = selection.Paragraphs[1].Range;
                }

                if (range == null || string.IsNullOrEmpty(range.Text))
                {
                    MessageBox.Show("未检测到可操作的文本或段落。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int rangeStart = range.Start;
                string originalText = range.Text;
                
                // 收集所有需要插入空格的位置（基于原始文本的相对位置）
                var insertPositions = new List<int>();
                
                // 复合单位列表（按长度从长到短排序，优先匹配长单位）
                string[] compoundUnits = { "kg/m³", "g/cm³", "kg/m²", "mol/L", "mol/l", "km/h", "m/s", "km/s", "Ω·m", "cm³", "cm²", "km³", "km²", "°C", "°F", "μm", "nm", "m³", "m²" };
                
                // 标记已被复合单位覆盖的位置
                var compoundPositions = new HashSet<int>();
                
                // 先匹配复合单位
                foreach (string unit in compoundUnits)
                {
                    string pattern = @"(\d+(?:\.\d+)?)(" + Regex.Escape(unit) + @")";
                    Regex regex = new Regex(pattern);
                    var matches = regex.Matches(originalText);
                    
                    foreach (Match match in matches)
                    {
                        int insertPos = match.Index + match.Groups[1].Length;
                        
                        // 检查边界和是否已有空格
                        if (insertPos >= 0 && insertPos < originalText.Length && originalText[insertPos] != ' ')
                        {
                            insertPositions.Add(insertPos);
                            
                            // 标记复合单位覆盖的所有字符位置
                            for (int i = insertPos; i < insertPos + unit.Length && i < originalText.Length; i++)
                            {
                                compoundPositions.Add(i);
                            }
                        }
                    }
                }
                
                // 匹配单个单位字符（排除百分比%、角度单位'和''，以及复合单位已覆盖的位置）
                string singleUnitPattern = @"(\d+(?:\.\d+)?)([a-zA-ZμΩℓ‰℃℉Å])";
                Regex singleUnitRegex = new Regex(singleUnitPattern);
                var singleMatches = singleUnitRegex.Matches(originalText);
                
                foreach (Match match in singleMatches)
                {
                    int insertPos = match.Index + match.Groups[1].Length;
                    
                    // 检查边界
                    if (insertPos < 0 || insertPos >= originalText.Length)
                    {
                        continue;
                    }
                    
                    // 如果已被复合单位覆盖，跳过
                    if (compoundPositions.Contains(insertPos))
                    {
                        continue;
                    }
                    
                    char nextChar = originalText[insertPos];
                    
                    // 排除百分比、角度单位
                    if (nextChar == '%' || nextChar == '\'' || nextChar == '"')
                    {
                        continue;
                    }
                    
                    // 如果已有空格，跳过
                    if (nextChar == ' ')
                    {
                        continue;
                    }
                    
                    insertPositions.Add(insertPos);
                }
                
                // 去重并排序，从后往前插入
                insertPositions = insertPositions.Distinct().OrderByDescending(p => p).ToList();
                
                // 从后往前插入空格
                foreach (int pos in insertPositions)
                {
                    Word.Range insertRange = range.Duplicate;
                    insertRange.Start = rangeStart + pos;
                    insertRange.End = insertRange.Start;
                    insertRange.InsertAfter(" ");
                }

                // 显示处理结果
                if (insertPositions.Count > 0)
                {
                    MessageBox.Show($"已为 {insertPositions.Count} 处数字和单位之间添加空格。", "完成", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("未找到需要添加空格的位置。", "提示", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"自动加空格失败：{ex.Message}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                Word.Range rng;

                if (sel != null && sel.Range != null && sel.Range.Start != sel.Range.End)
                {
                    rng = sel.Range;
                }
                else
                {
                    rng = app.ActiveDocument.Range();
                }

                rng.Find.ClearFormatting();
                rng.Find.Replacement.ClearFormatting();
                
                // 设置查找字体（原字体可能不存在，但文档中已使用）
                try
                {
                    rng.Find.Font.Name = originalFont;
                }
                catch
                {
                    // 原字体设置失败不影响，继续
                }
                
                // 设置替换字体（新字体需要存在）
                try
                {
                    rng.Find.Replacement.Font.Name = newFont;
                }
                catch
                {
                    MessageBox.Show($"无法设置替换字体 '{newFont}'，该字体可能不存在。\n\n请确保系统中已安装该字体。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
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
            catch (Exception ex)
            {
                MessageBox.Show($"字体替换失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ReplaceFangSongGB2312ToFangSong()
        {
            ReplaceFont("仿宋_GB2312", "仿宋");
        }

        public void ReplaceKaiTiGB2312ToKaiTi()
        {
            ReplaceFont("楷体_GB2312", "楷体");
        }

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