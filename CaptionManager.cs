using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan
{
    public class CaptionManager
    {
        #region 枚举定义
        public enum CaptionNumberStyle
        {
            Arabic,     // 图 1 / 表 1
            Dash,       // 图 1-1 / 表 1-1
            Dot         // 图 1.1 / 表 1.1
        }

        public enum FormulaNumberStyle
        {
            Parenthesis1,    // (1)
            Parenthesis1_1,  // (1-1)
            Parenthesis1_1dot// (1.1)
        }
        #endregion

        #region 状态管理
        private FormulaNumberStyle currentFormulaStyle = FormulaNumberStyle.Parenthesis1;
        private CaptionNumberStyle currentPictureStyle = CaptionNumberStyle.Arabic;
        private CaptionNumberStyle currentTableStyle = CaptionNumberStyle.Arabic;

        // 交叉引用相关字段
        private Word.Range originalRange;
        private bool isCrossReferenceMode = false;
        private RibbonToggleButton crossReferenceToggleButton;
        private System.Windows.Forms.Timer escKeyListener;

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(Keys vKey);

        private void SetStyle(RibbonToggleButton selected, RibbonToggleButton other1, RibbonToggleButton other2)
        {
            selected.Checked = true;
            other1.Checked = false;
            other2.Checked = false;
        }
        #endregion

        #region 式编号相关方法
        public void SetFormulaStyle(RibbonToggleButton selected, RibbonToggleButton other1, RibbonToggleButton other2, FormulaNumberStyle style)
        {
            SetStyle(selected, other1, other2);
            currentFormulaStyle = style;
        }

        public void InsertFormulaNumber()
        {
            int originalStart = 0;
            int originalEnd = 0;

            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                var doc = app.ActiveDocument;

                if (doc == null)
                {
                    MessageBox.Show("未检测到活动文档。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (sel == null || sel.Paragraphs == null || sel.Paragraphs.Count == 0)
                {
                    MessageBox.Show("请将光标放在包含公式的段落中。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                originalStart = sel.Start;
                originalEnd = sel.End;

                Word.Paragraph para = sel.Paragraphs[1];
                if (para == null || para.Range == null)
                {
                    MessageBox.Show("无法获取当前段落。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 检查前一个段落是否为空
                Word.Paragraph prevPara = para.Previous() as Word.Paragraph;
                bool prevParaIsEmpty = false;
                if (prevPara != null)
                {
                    string prevText = prevPara.Range.Text.Trim();
                    prevParaIsEmpty = string.IsNullOrEmpty(prevText) || prevText == "\r" || prevText == "\r\n";
                }

                // 如果前一个段落为空，在删除当前段落前，先在前一个段落末尾插入一个段落
                // 这样可以确保即使当前段落被删除，前一个空段落也会保留
                if (prevParaIsEmpty && prevPara != null)
                {
                    Word.Range prevParaEnd = prevPara.Range.Duplicate;
                    prevParaEnd.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    prevParaEnd.InsertParagraphAfter();
                }

                Word.Range contentRange = para.Range.Duplicate;
                contentRange.End = contentRange.End - 1;
                contentRange.Cut();

                // 删除当前段落（包括段落标记）
                para.Range.Delete();

                Word.Table table = CreateFormulaTable(sel, app);
                if (table == null)
                {
                    MessageBox.Show("创建公式表格失败。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                table.Cell(1, 2).Range.Paste();

                InsertFormulaNumber(table, sel, currentFormulaStyle);

                table.Cell(1, 2).Range.Select();
            }
            catch (Exception ex)
            {
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    var sel = app.Selection;
                    if (sel != null && originalStart > 0 && originalEnd > 0)
                    {
                        sel.SetRange(originalStart, originalEnd);
                    }
                }
                catch (Exception restoreEx)
                {
                    System.Diagnostics.Debug.WriteLine($"恢复选择位置失败: {restoreEx.Message}");
                }

                MessageBox.Show($"式编号插入失败：{ex.Message}\n\n请确保光标位于包含公式的段落中。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Word.Table CreateFormulaTable(Word.Selection sel, Word.Application app)
        {
            Word.Table table = sel.Tables.Add(sel.Range, 1, 3);

            table.Borders.Enable = 0;

            float pageWidth = sel.PageSetup.PageWidth - sel.PageSetup.LeftMargin - sel.PageSetup.RightMargin;
            float[] columnWidths = { pageWidth * 0.15f, pageWidth * 0.7f, pageWidth * 0.15f };

            for (int i = 0; i < 3; i++)
            {
                table.Columns[i + 1].Width = app.CentimetersToPoints(columnWidths[i] / 28.35f);
            }

            Word.WdParagraphAlignment[] alignments =
            {
                Word.WdParagraphAlignment.wdAlignParagraphLeft,
                Word.WdParagraphAlignment.wdAlignParagraphCenter,
                Word.WdParagraphAlignment.wdAlignParagraphRight
            };

            for (int i = 0; i < 3; i++)
            {
                table.Cell(1, i + 1).Range.ParagraphFormat.Alignment = alignments[i];
            }

            return table;
        }

        private void InsertFormulaNumber(Word.Table table, Word.Selection sel, FormulaNumberStyle currentStyle)
        {
            const string leftBracket = "(";
            const string rightBracket = ")";
            const string seqName = "公式";

            table.Cell(1, 3).Range.Select();
            sel.TypeText(leftBracket);

            switch (currentStyle)
            {
                case FormulaNumberStyle.Parenthesis1:
                    sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldSequence, seqName, false);
                    break;

                case FormulaNumberStyle.Parenthesis1_1:
                    sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
                    sel.TypeText("-");
                    sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldSequence, seqName + "\\s 1", false);
                    break;

                case FormulaNumberStyle.Parenthesis1_1dot:
                    sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
                    sel.TypeText(".");
                    sel.Fields.Add(sel.Range, Word.WdFieldType.wdFieldSequence, seqName + "\\s 1", false);
                    break;
            }

            sel.TypeText(rightBracket);
        }
        #endregion

        #region 图编号相关方法
        public void SetPictureStyle(RibbonToggleButton selected, RibbonToggleButton other1, RibbonToggleButton other2, CaptionNumberStyle style)
        {
            SetStyle(selected, other1, other2);
            currentPictureStyle = style;
        }

        public void InsertPictureNumber()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                var doc = app.ActiveDocument;

                if (!ValidateDocumentAndSelection(doc, sel))
                {
                    return;
                }

                app.ScreenUpdating = false;

                HashSet<int> handledParagraphs = new HashSet<int>();
                List<Word.Paragraph> targetParas = new List<Word.Paragraph>();

                try
                {
                    if (sel.Range.InlineShapes != null && sel.Range.InlineShapes.Count > 0)
                    {
                        foreach (Word.InlineShape ils in sel.Range.InlineShapes)
                        {
                            if (ils.Range != null && ils.Range.Paragraphs != null && ils.Range.Paragraphs.Count > 0)
                            {
                                var para = ils.Range.Paragraphs[1];
                                if (para != null && para.Range != null && !handledParagraphs.Contains(para.Range.Start))
                                {
                                    targetParas.Add(para);
                                    handledParagraphs.Add(para.Range.Start);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"处理内嵌图片时出错: {ex.Message}");
                }

                try
                {
                    if (sel.Range.ShapeRange != null && sel.Range.ShapeRange.Count > 0)
                    {
                        foreach (Word.Shape s in sel.Range.ShapeRange)
                        {
                            if (s.Anchor != null && s.Anchor.Paragraphs != null && s.Anchor.Paragraphs.Count > 0)
                            {
                                var para = s.Anchor.Paragraphs[1];
                                if (para != null && para.Range != null && !handledParagraphs.Contains(para.Range.Start))
                                {
                                    targetParas.Add(para);
                                    handledParagraphs.Add(para.Range.Start);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"处理浮动图片时出错: {ex.Message}");
                }

                if (targetParas.Count == 0 && sel.Paragraphs != null && sel.Paragraphs.Count > 0)
                {
                    var para = sel.Paragraphs[1];
                    if (para != null && para.Range != null && !handledParagraphs.Contains(para.Range.Start))
                    {
                        targetParas.Add(para);
                        handledParagraphs.Add(para.Range.Start);
                    }
                }

                if (targetParas.Count == 0)
                {
                    MessageBox.Show("未找到可插入题注的图片。\n\n请确保：\n1. 光标位于包含图片的段落中\n2. 或选中了包含图片的内容", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                for (int i = targetParas.Count - 1; i >= 0; i--)
                {
                    InsertPictureCaption(targetParas[i], currentPictureStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"图编号插入失败：{ex.Message}\n\n请确保光标位于包含图片的段落中。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }

        private void InsertPictureCaption(Word.Paragraph picPara, CaptionNumberStyle numberStyle)
        {
            if (picPara == null) return;

            var doc = picPara.Range.Application.ActiveDocument;

            // 保存图片段落结束位置
            int originalPicEnd = picPara.Range.End;
            
            // 检查图片段落后是否已经有段落
            var nextPara = picPara.Next() as Word.Paragraph;
            Word.Paragraph captionPara = null;
            
            if (nextPara != null)
            {
                string nextText = nextPara.Range.Text.Trim();
                Word.Style nextStyle = null;
                try
                {
                    nextStyle = nextPara.get_Style() as Word.Style;
                }
                catch { }
                
                // 如果下一个段落已经是题注样式，或者已经包含"图"开头的文本，直接返回
                if ((nextStyle != null && nextStyle.NameLocal == "题注") || nextText.StartsWith("图"))
                    {
                        return;
                    }
                
                // 如果下一个段落紧挨着图片段落（开始位置等于图片段落结束位置），且是空段落，可以使用
                if (nextPara.Range.Start == originalPicEnd && string.IsNullOrEmpty(nextText))
                {
                    captionPara = nextPara;
                }
            }
            
            // 如果没有找到可用的空段落，需要插入新段落
            if (captionPara == null)
            {
                // 如果后面有文本段落，在文本段落之前插入新段落，避免合并
                if (nextPara != null && !string.IsNullOrEmpty(nextPara.Range.Text.Trim()))
                {
                    int originalTextParaStart = nextPara.Range.Start;
                    Word.Range beforeTextPara = doc.Range(originalTextParaStart, originalTextParaStart);
                    beforeTextPara.InsertParagraphBefore();
                    
                    // 查找新插入的段落（插入后，新段落的开始位置应该是原文本段落的开始位置）
                    foreach (Word.Paragraph para in doc.Paragraphs)
                    {
                        if (para.Range.Start == originalTextParaStart)
                        {
                            captionPara = para;
                            break;
                        }
                    }
                }
                else
                {
                    // 图片后面没有文本，直接在图片段落后插入新段落
            var afterPicRange = picPara.Range.Duplicate;
            afterPicRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            afterPicRange.InsertParagraphAfter();

                    // 查找新插入的题注段落
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                        if (para.Range.Start == originalPicEnd)
                {
                    captionPara = para;
                    break;
                        }
                    }
                }
            }

            if (captionPara == null) return;

            // 清除段落内容，准备插入题注
            Word.Range captionRange = captionPara.Range.Duplicate;
            captionRange.End = captionRange.End - 1; // 排除段落标记
            captionRange.Text = "";

            // 在题注段落中插入内容
            captionRange = captionPara.Range.Duplicate;
            captionRange.End = captionRange.End - 1; // 排除段落标记
            captionRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            
            captionRange.InsertAfter("图 ");
            captionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            switch (numberStyle)
            {
                case CaptionNumberStyle.Arabic:
                    captionRange.Fields.Add(captionRange, Word.WdFieldType.wdFieldSequence, "图 \\* ARABIC", false);
                    break;

                case CaptionNumberStyle.Dash:
                case CaptionNumberStyle.Dot:
                    InsertNumberWithStyleRef(captionRange, "图", numberStyle == CaptionNumberStyle.Dash ? "-" : ".");
                    break;
            }

            captionPara.set_Style("题注");
        }
        #endregion

        #region 表编号相关方法
        public void SetTableStyle(RibbonToggleButton selected, RibbonToggleButton other1, RibbonToggleButton other2, CaptionNumberStyle style)
        {
            SetStyle(selected, other1, other2);
            currentTableStyle = style;
        }

        public void InsertTableNumber()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                var doc = app.ActiveDocument;

                if (!ValidateDocumentAndSelection(doc, sel))
                {
                    return;
                }

                app.ScreenUpdating = false;

                HashSet<int> handledTables = new HashSet<int>();
                List<Word.Table> targetTables = new List<Word.Table>();

                if (doc.Tables != null && doc.Tables.Count > 0)
                {
                    foreach (Word.Table table in doc.Tables)
                    {
                        try
                        {
                            if (table != null && table.Range != null && sel.Range != null)
                            {
                                if (table.Range.InRange(sel.Range) && !handledTables.Contains(table.Range.Start))
                                {
                                    targetTables.Add(table);
                                    handledTables.Add(table.Range.Start);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"处理表格时出错: {ex.Message}");
                        }
                    }
                }

                if (targetTables.Count == 0 && sel.Tables != null && sel.Tables.Count > 0)
                {
                    var table = sel.Tables[1];
                    if (table != null && table.Range != null && !handledTables.Contains(table.Range.Start))
                    {
                        targetTables.Add(table);
                        handledTables.Add(table.Range.Start);
                    }
                }

                if (targetTables.Count == 0)
                {
                    MessageBox.Show("未找到可插入题注的表格。\n\n请确保：\n1. 光标位于表格中\n2. 或选中了表格内容", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                for (int i = targetTables.Count - 1; i >= 0; i--)
                {
                    InsertTableCaption(targetTables[i], currentTableStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"表编号插入失败：{ex.Message}\n\n请确保光标位于包含表格的段落中。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }

        private void InsertTableCaption(Word.Table table, CaptionNumberStyle numberStyle)
        {
            if (table == null) return;
            
            try
            {
                var doc = table.Application.ActiveDocument;
                var app = doc.Application;

                // 检查表格前是否已有题注
                if (HasCaptionBeforeTable(table, doc))
                {
                    return;
                }

                // 1. 定位到表格的第一个单元格
                Word.Range firstCellRange = table.Cell(1, 1).Range;
                app.Selection.SetRange(firstCellRange.Start, firstCellRange.Start);

                // 2. 按上键，检查位置是否还在第一个单元格
                SendKeyAndWait("{UP}");
                bool isInFirstCell = table.Cell(1, 1).Range.InRange(app.Selection.Range);

                if (isInFirstCell)
                {
                    // 位置还在第一个单元格，说明前方没有段落
                    // 使用 Ctrl+Shift+Enter 在表格前插入段落
                    SendKeyAndWait("^+{ENTER}");
                }
                else
                {
                    // 位置不在第一个单元格了，说明前方有段落
                    // 判断当前光标所在段落是否有内容
                    Word.Paragraph currentPara = app.Selection.Paragraphs[1];
                    if (currentPara != null && IsParagraphNotEmpty(currentPara))
                    {
                        // 段落有内容，在段落最后插入一个空行
                        app.Selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);
                        SendKeyAndWait("{ENTER}");
                    }
                    // 如果段落没有内容（是空段落），直接使用，不需要插入
                }

                // 3. 将光标定位到空段落开始位置
                app.Selection.HomeKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

                // 4. 插入编号
                InsertTableCaptionContent(app.Selection.Range, doc, numberStyle);

                // 5. 应用题注样式
                Word.Paragraph captionPara = app.Selection.Paragraphs[1];
                if (captionPara != null)
                {
                    captionPara.set_Style("题注");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入表编号失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool HasCaptionBeforeTable(Word.Table table, Word.Document doc)
        {
            Word.Range tableRange = table.Range;
            Word.Range beforeTableRange = doc.Range(0, tableRange.Start);
            if (beforeTableRange.Paragraphs.Count > 0)
            {
                Word.Paragraph prevPara = beforeTableRange.Paragraphs[beforeTableRange.Paragraphs.Count];
                string prevText = prevPara.Range.Text.TrimStart();
                Word.Style prevStyle = prevPara.get_Style() as Word.Style;
                
                return (prevStyle != null && prevStyle.NameLocal == "题注") || prevText.StartsWith("表");
            }
            return false;
        }

        private bool IsParagraphNotEmpty(Word.Paragraph para)
        {
            if (para == null) return false;
            string paraText = para.Range.Text.Trim();
            return !string.IsNullOrEmpty(paraText) && paraText != "\r" && paraText != "\r\n";
        }

        private void SendKeyAndWait(string keys)
        {
            System.Windows.Forms.SendKeys.SendWait(keys);
            System.Threading.Thread.Sleep(50);
        }

        private void InsertTableCaptionContent(Word.Range range, Word.Document doc, CaptionNumberStyle numberStyle)
        {
            Word.Range captionRange = range.Duplicate;
            captionRange.Text = "";
            var fieldRange = doc.Range(captionRange.Start, captionRange.Start);
            fieldRange.InsertAfter("表 ");
            fieldRange.SetRange(fieldRange.Start + 2, fieldRange.Start + 2);

            switch (numberStyle)
            {
                case CaptionNumberStyle.Arabic:
                    fieldRange.Fields.Add(fieldRange, Word.WdFieldType.wdFieldSequence, "表 \\* ARABIC", false);
                    break;
                case CaptionNumberStyle.Dash:
                case CaptionNumberStyle.Dot:
                    InsertNumberWithStyleRef(fieldRange, "表", numberStyle == CaptionNumberStyle.Dash ? "-" : ".");
                    break;
            }
        }
        #endregion

        #region 交叉引用相关方法
        public class CaptionInfo
        {
            public string Identifier { get; set; }
            public string Number { get; set; }
            public string FullText { get; set; }
        }

        public void ToggleCrossReferenceMode(RibbonToggleButton toggleButton)
        {
            crossReferenceToggleButton = toggleButton;

            if (isCrossReferenceMode)
            {
                ExitCrossReferenceMode();
                return;
            }

            try
            {
                originalRange = Globals.ThisAddIn.Application.Selection.Range;
                isCrossReferenceMode = true;
                crossReferenceToggleButton.Checked = true;

                Globals.ThisAddIn.Application.WindowSelectionChange += Application_WindowSelectionChange;
                InitializeEscKeyListener();

                Globals.ThisAddIn.Application.StatusBar = "交叉引用模式：请将光标移动到题注所在行，按ESC或再次点击按钮退出";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动交叉引用模式失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                isCrossReferenceMode = false;
                crossReferenceToggleButton.Checked = false;
            }
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            if (!isCrossReferenceMode) return;

            Word.Range currentRange = Sel.Range;
            CaptionInfo captionInfo = FindCaptionInfo(currentRange);

            if (captionInfo != null)
            {
                InsertCrossReferenceAtOriginalPosition(originalRange, captionInfo);
                ExitCrossReferenceMode();
            }
        }

        private void InitializeEscKeyListener()
        {
            escKeyListener = new System.Windows.Forms.Timer();
            escKeyListener.Interval = 100;
            escKeyListener.Tick += EscKeyListener_Tick;
            escKeyListener.Start();
        }

        private void EscKeyListener_Tick(object sender, EventArgs e)
        {
            if (isCrossReferenceMode && (GetAsyncKeyState(Keys.Escape) & 0x8000) != 0)
            {
                ExitCrossReferenceMode();
            }
        }

        private void ExitCrossReferenceMode()
        {
            isCrossReferenceMode = false;
            Globals.ThisAddIn.Application.WindowSelectionChange -= Application_WindowSelectionChange;

            if (escKeyListener != null)
            {
                escKeyListener.Stop();
                escKeyListener.Dispose();
                escKeyListener = null;
            }

            Globals.ThisAddIn.Application.StatusBar = "";

            if (crossReferenceToggleButton != null)
            {
                crossReferenceToggleButton.Checked = false;
            }

            originalRange = null;
        }

        private CaptionInfo FindCaptionInfo(Word.Range range)
        {
            if (range == null || range.Paragraphs == null || range.Paragraphs.Count == 0)
            {
                return null;
            }

            Word.Paragraph paragraph = range.Paragraphs[1];
            if (paragraph == null || paragraph.Range == null)
            {
                return null;
            }

            Word.Range paraRange = paragraph.Range;

            foreach (Word.Field field in paraRange.Fields)
            {
                if (field != null && field.Type == Word.WdFieldType.wdFieldSequence)
                {
                    string fieldCode = field.Code != null ? field.Code.Text : string.Empty;
                    string identifier = ExtractIdentifierFromFieldCode(fieldCode);

                    string captionText = field.Result != null ? field.Result.Text.Trim() : string.Empty;

                    string number = ExtractNumberFromCaption(captionText);

                    if (!string.IsNullOrEmpty(identifier) && !string.IsNullOrEmpty(number))
                    {
                        return new CaptionInfo
                        {
                            Identifier = identifier,
                            Number = number,
                            FullText = captionText
                        };
                    }
                }
            }

            return null;
        }

        private string ExtractIdentifierFromFieldCode(string fieldCode)
        {
            fieldCode = fieldCode.Trim();

            if (fieldCode.StartsWith("SEQ", StringComparison.OrdinalIgnoreCase))
            {
                string remaining = fieldCode.Substring(3).Trim();

                int spaceIndex = remaining.IndexOf(' ');
                int backslashIndex = remaining.IndexOf('\\');

                int endIndex = -1;
                if (spaceIndex >= 0 && backslashIndex >= 0)
                    endIndex = Math.Min(spaceIndex, backslashIndex);
                else if (spaceIndex >= 0)
                    endIndex = spaceIndex;
                else if (backslashIndex >= 0)
                    endIndex = backslashIndex;

                if (endIndex >= 0)
                    return remaining.Substring(0, endIndex).Trim();
                else
                    return remaining.Trim();
            }

            return string.Empty;
        }

        private string ExtractNumberFromCaption(string captionText)
        {
            if (string.IsNullOrEmpty(captionText))
                return string.Empty;

            string[] prefixes = { "图 ", "表 ", "公式 " };
            foreach (string prefix in prefixes)
            {
                if (captionText.StartsWith(prefix))
                {
                    return captionText.Substring(prefix.Length).Trim();
                }
            }

            return captionText.Trim();
        }

        private void InsertCrossReferenceAtOriginalPosition(Word.Range originalRange, CaptionInfo captionInfo)
        {
            Globals.ThisAddIn.Application.Selection.SetRange(
                originalRange.Start, originalRange.End);

            object referenceType = captionInfo.Identifier;

            Word.WdReferenceKind referenceKind = (captionInfo.Identifier == "公式" || captionInfo.Identifier == "式" || captionInfo.Identifier == "EQ") ?
            Word.WdReferenceKind.wdEntireCaption : Word.WdReferenceKind.wdOnlyLabelAndNumber;

            object referenceItem = captionInfo.Number;

            Globals.ThisAddIn.Application.Selection.InsertCrossReference(
                referenceType,
                referenceKind,
                referenceItem,
                System.Type.Missing,
                System.Type.Missing,
                System.Type.Missing,
                System.Type.Missing);
        }
        #endregion

        #region 辅助方法
        private bool ValidateDocumentAndSelection(Word.Document doc, Word.Selection sel)
        {
            if (doc == null)
            {
                MessageBox.Show("未检测到活动文档。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (sel == null)
            {
                MessageBox.Show("无法获取当前选择。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void InsertNumberWithStyleRef(Word.Range range, string seqName, string separator)
        {
            var styleRefField = range.Fields.Add(
                range, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
            styleRefField.Result.Select();
            var selection = range.Application.Selection;
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

            selection.TypeText(separator);

            selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

            selection.Range.Fields.Add(
                selection.Range, Word.WdFieldType.wdFieldSequence, seqName + " \\s 1", false);
        }
        #endregion
    }
}
