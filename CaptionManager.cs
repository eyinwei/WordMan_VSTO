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

                Word.Range contentRange = para.Range.Duplicate;
                contentRange.End = contentRange.End - 1;
                contentRange.Cut();

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

            var nextPara = picPara.Next() as Word.Paragraph;
            if (nextPara != null)
            {
                string nextText = nextPara.Range.Text.Trim();
                if (!string.IsNullOrEmpty(nextText))
                {
                    if ((nextPara.get_Style() is Word.Style style && style.NameLocal == "题注")
                        || nextText.StartsWith("图"))
                    {
                        return;
                    }
                }
            }

            int originalPicPosition = picPara.Range.End;

            var afterPicRange = picPara.Range.Duplicate;
            afterPicRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            afterPicRange.InsertParagraphAfter();

            Word.Paragraph captionPara = null;
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                if (para.Range.Start == originalPicPosition)
                {
                    captionPara = para;
                    break;
                }
            }

            if (captionPara == null) return;

            Word.Range captionRange = captionPara.Range.Duplicate;
            captionRange.End = captionRange.Start + 1;
            captionRange.Text = "";

            var insertRange = doc.Range(captionRange.Start, captionRange.Start);
            insertRange.InsertAfter("图 ");
            insertRange.SetRange(insertRange.Start + 2, insertRange.Start + 2);

            switch (numberStyle)
            {
                case CaptionNumberStyle.Arabic:
                    insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldSequence, "图 \\* ARABIC", false);
                    break;

                case CaptionNumberStyle.Dash:
                case CaptionNumberStyle.Dot:
                    InsertNumberWithStyleRef(insertRange, "图", numberStyle == CaptionNumberStyle.Dash ? "-" : ".");
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
            var doc = table.Application.ActiveDocument;
            var app = doc.Application;

            Word.Range tableRange = table.Range;
            int tableStart = tableRange.Start;

            bool isInFirstCell = table.Cell(1, 1).Range.InRange(tableRange);
            if (isInFirstCell)
            {
                tableStart = Math.Max(0, tableStart - 1);
            }

            Word.Paragraph prevPara = null;
            Word.Range beforeTableRange = doc.Range(0, tableStart);
            if (beforeTableRange.Paragraphs.Count > 0)
            {
                prevPara = beforeTableRange.Paragraphs[beforeTableRange.Paragraphs.Count];
                string prevText = prevPara.Range.Text.TrimStart();
                if ((prevPara.get_Style() is Word.Style style && style.NameLocal == "题注")
                    || prevText.StartsWith("表"))
                {
                    return;
                }
            }

            int originalTablePosition = tableRange.Start;

            Word.Range insertRange = doc.Range(tableStart, tableStart);
            insertRange.Text = "";
            insertRange.InsertParagraphBefore();

            Word.Paragraph captionPara = null;
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                if (para.Range.End == originalTablePosition)
                {
                    captionPara = para;
                    break;
                }
            }

            if (captionPara == null) return;

            Word.Range captionRange = captionPara.Range.Duplicate;
            captionRange.End = captionRange.Start + 1;
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

            captionPara.set_Style("题注");
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
