using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Microsoft.VisualBasic;    
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordMan.SplitAndMerge;
using WordMan.MultiLevel;

namespace WordMan
{
    public partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {                                   

        }
        private void 去除断行_Click(object sender, RibbonControlEventArgs e)
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
        private void 去除空格_Click(object sender, RibbonControlEventArgs e)
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
        private void 去除空行_Click(object sender, RibbonControlEventArgs e)
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
        private void 去除缩进_Click(object sender, RibbonControlEventArgs e)
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
        private void 缩进2字符_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.Selection;
            if (selection != null)
            {
                var paraFormat = selection.ParagraphFormat;
                paraFormat.CharacterUnitFirstLineIndent = 2f;
            }
        }
        private void 希腊字母_Click(object sender, RibbonControlEventArgs e)
        {
            GreekLetterForm form = new GreekLetterForm();
            form.Show();
        }
        private void 常用符号_Click(object sender, RibbonControlEventArgs e)
        {
            CommonSymbolForm form = new CommonSymbolForm();
            form.Show();
        }
        private void 另存PDF_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;

            // 1. 检查文档是否已保存过
            if (string.IsNullOrEmpty(doc.Path))
            {
                System.Windows.Forms.MessageBox.Show(
                    "请先保存文档，再导出为PDF。",
                    "提示",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);

                // 调用Word的“另存为”对话框
                app.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFileSaveAs].Show();

                // 不再自动导出PDF，无论保存没保存，直接退出
                return;
            }

            try
            {
                string docPath = doc.FullName;
                string directory = System.IO.Path.GetDirectoryName(docPath);
                string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(docPath);
                string pdfPath = System.IO.Path.Combine(directory, fileNameWithoutExt + ".pdf");

                // 2. 导出为PDF，设置 OpenAfterExport 为 false
                doc.ExportAsFixedFormat(
                    pdfPath,
                    Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false, // 不自动打开PDF
                    OptimizeFor: Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range: Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument,
                    Item: Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent,
                    CreateBookmarks: Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                    DocStructureTags: true,
                    BitmapMissingFonts: true,
                    UseISO19005_1: false
                );

                // 3. 成功后弹窗，询问是否打开PDF
                var result = System.Windows.Forms.MessageBox.Show(
                    "成功导出为PDF！是否现在打开该PDF？",
                    "导出成功",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question);

                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(pdfPath);
                    }
                    catch (Exception exOpen)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "打开PDF文件出错：" + exOpen.Message,
                            "错误",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "导出PDF失败：" + ex.Message,
                    "错误",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        private void 英中标点互转_Click(object sender, RibbonControlEventArgs e, bool englishToChinese)
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
                foreach (var pair in dict)
                {
                    try
                    {
                        rng.Find.ClearFormatting();
                        rng.Find.Text = pair.Key;
                        rng.Find.Replacement.ClearFormatting();
                        rng.Find.Replacement.Text = pair.Value;
                        rng.Find.MatchWildcards = false;
                        rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
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
                        
                        // 处理双引号 - 逐个替换
                        bool isLeft = true;
                        var searchRange = app.ActiveDocument.Range(startPos, endPos);
                        searchRange.Find.ClearFormatting();
                        searchRange.Find.Text = "\"";
                        searchRange.Find.Forward = true;
                        searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;
                        
                        while (searchRange.Find.Execute())
                        {
                            var hitRange = app.ActiveDocument.Range(searchRange.Start, searchRange.Start + 1);
                            hitRange.Text = isLeft ? "\u201c" : "\u201d";
                            isLeft = !isLeft;
                            // 更新搜索范围，从当前位置继续
                            searchRange.SetRange(hitRange.End, endPos);
                        }
                        
                        // 处理单引号 - 逐个替换
                        isLeft = true;
                        searchRange = app.ActiveDocument.Range(startPos, endPos);
                        searchRange.Find.ClearFormatting();
                        searchRange.Find.Text = "'";
                        searchRange.Find.Forward = true;
                        searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;
                        
                        while (searchRange.Find.Execute())
                        {
                            var hitRange = app.ActiveDocument.Range(searchRange.Start, searchRange.Start + 1);
                            hitRange.Text = isLeft ? "\u2018" : "\u2019";
                            isLeft = !isLeft;
                            // 更新搜索范围，从当前位置继续
                            searchRange.SetRange(hitRange.End, endPos);
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
            
        private void 英标转中标_Click(object sender, RibbonControlEventArgs e)
        {
            英中标点互转_Click(sender, e, true);
        }
        private void 中标转英标_Click(object sender, RibbonControlEventArgs e)
        {
            英中标点互转_Click(sender, e, false);
        }

        private void 自动加空格_Click(object sender, RibbonControlEventArgs e)
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

        public enum FormulaNumberStyle
        {
            Parenthesis1,    // (1)
            Parenthesis1_1,  // (1-1)
            Parenthesis1_1dot// (1.1)
        }
        private FormulaNumberStyle CurrentStyle = FormulaNumberStyle.Parenthesis1;

        private void 公式样式1_Click(object sender, RibbonControlEventArgs e)
        {
            公式样式1.Checked = true;
            公式样式2.Checked = false;
            公式样式3.Checked = false;
            CurrentStyle = FormulaNumberStyle.Parenthesis1;
        }
        private void 公式样式2_Click(object sender, RibbonControlEventArgs e)
        {
            公式样式1.Checked = false;
            公式样式2.Checked = true;
            公式样式3.Checked = false;
            CurrentStyle = FormulaNumberStyle.Parenthesis1_1;
        }
        private void 公式样式3_Click(object sender, RibbonControlEventArgs e)
        {
            公式样式1.Checked = false;
            公式样式2.Checked = false;
            公式样式3.Checked = true;
            CurrentStyle = FormulaNumberStyle.Parenthesis1_1dot;
        }

        private void 公式编号_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            Word.Paragraph para = sel.Paragraphs[1];

            // 1. 在该行Home插入一个居中制表位，在End插入一个居右制表位
            float pageWidth = sel.PageSetup.PageWidth - sel.PageSetup.LeftMargin - sel.PageSetup.RightMargin;
            float centerPos = pageWidth / 2.0f;
            float rightPos = pageWidth;
            para.TabStops.ClearAll();
            para.TabStops.Add(app.CentimetersToPoints(centerPos / 28.35f),
                              Word.WdTabAlignment.wdAlignTabCenter,
                              Word.WdTabLeader.wdTabLeaderSpaces);
            para.TabStops.Add(app.CentimetersToPoints(rightPos / 28.35f),
                              Word.WdTabAlignment.wdAlignTabRight,
                              Word.WdTabLeader.wdTabLeaderSpaces);

            // 2. 段首插入Tab（连续两次HomeKey）
            sel.SetRange(para.Range.Start, para.Range.Start);
            sel.HomeKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);
            sel.HomeKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);
            sel.TypeText("\t");

            // 段尾插入Tab（连续两次EndKey）
            sel.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);
            sel.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);
            sel.TypeText("\t");


            // 3. 在尾Tab后插入括号和编号

            // 获取尾Tab后的Range
            Word.Range insertRange = para.Range.Duplicate;
            insertRange.Start = insertRange.End - 1;
            insertRange.End = insertRange.End - 1;

            // 括号风格
            string leftBracket = "(";
            string rightBracket = ")";
            string seqName = "公式";

            insertRange.InsertAfter(leftBracket);
            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            switch (CurrentStyle)
            {
                case FormulaNumberStyle.Parenthesis1:
                    // 公式（1） ==> 1 由 SEQ 公式 得到
                    var seqField = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldSequence, seqName, false);
                    // 确保移动到域的结束位置之后
                    insertRange.Move(Word.WdUnits.wdCharacter, seqField.Result.Characters.Count);
                    break;

                case FormulaNumberStyle.Parenthesis1_1:
                    var srField2 = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
                    insertRange.Move(Word.WdUnits.wdCharacter, srField2.Result.Characters.Count);

                    insertRange.InsertAfter("-");
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    var seqField2 = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldSequence, seqName + "\\s 1", false);
                    insertRange.Move(Word.WdUnits.wdCharacter, seqField2.Result.Characters.Count);
                    break;

                case FormulaNumberStyle.Parenthesis1_1dot:
                    var srField3 = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
                    insertRange.Move(Word.WdUnits.wdCharacter, srField3.Result.Characters.Count);

                    insertRange.InsertAfter(".");
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    var seqField3 = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldSequence, seqName + "\\s 1", false);
                    insertRange.Move(Word.WdUnits.wdCharacter, seqField3.Result.Characters.Count);
                    break;
            }

            insertRange.InsertAfter(rightBracket);
        }
        private void 创建三线表_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;

            // 1. 创建3x2表格
            Word.Table table = sel.Tables.Add(sel.Range, 3, 3);

            // 2. 选中整个表格
            table.Select();

            // 3. 调用已有的设为三线表方法
            设为三线表_Click(sender, e);

        }

        private void 设为三线表_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
    {
        var app = Globals.ThisAddIn.Application;
        var sel = app.Selection;
        if (sel == null || sel.Tables.Count == 0)
            return;

        Word.Table table = sel.Tables[1];
        int firstRowIndex = int.MaxValue;
        int lastRowIndex = int.MinValue;

        // 首先找出最小和最大行号（因为有合并单元格，不能用Rows.Count）
        foreach (Word.Cell cell in table.Range.Cells)
        {
            if (cell.RowIndex < firstRowIndex)
                firstRowIndex = cell.RowIndex;
            if (cell.RowIndex > lastRowIndex)
                lastRowIndex = cell.RowIndex;
        }

        // 遍历所有单元格，清除所有边框
        foreach (Word.Cell cell in table.Range.Cells)
        {
            foreach (Word.WdBorderType borderType in new[]
            {
            Word.WdBorderType.wdBorderLeft,
            Word.WdBorderType.wdBorderRight,
            Word.WdBorderType.wdBorderTop,
            Word.WdBorderType.wdBorderBottom
        })
            {
                cell.Borders[borderType].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            }
        }

        // 遍历所有单元格，为三线表加主线
        foreach (Word.Cell cell in table.Range.Cells)
        {
            if (cell.RowIndex == firstRowIndex)
            {
                // 第一行：加上边粗线
                cell.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                cell.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth150pt; // 1.5磅

                // 第一行：加下边细线（即三线表“中线”）
                cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                cell.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth075pt; // 0.75磅
            }
            if (cell.RowIndex == lastRowIndex)
            {
                // 最后一行：加下边粗线
                cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                cell.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt; // 1.5磅
            }
        }

        // 可以额外设置格式和对齐等
        table.Range.Font.Size = 10.5f;
        table.Range.Font.NameFarEast = "宋体";
        table.Range.Font.Name = "Times New Roman";
        table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
        table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        table.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

        // 自动适应
        table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
        table.PreferredWidth = 100f;
        table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
    }


        private void 插入N行_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;

            if (sel == null || sel.Tables.Count == 0)
            {
                MessageBox.Show("请将光标放在表格内！", "提示");
                return;
            }

            // 输入要插入的行数
            string input = Interaction.InputBox("请输入要插入的行数：", "插入行", "1");
            if (string.IsNullOrWhiteSpace(input))
                return;

            if (!int.TryParse(input, out int n) || n <= 0)
            {
                MessageBox.Show("请输入有效的正整数！", "提示");
                return;
            }

            // 选择插入方向
            var direction = MessageBox.Show(
                "点击“是”在上方插入，点击“否”在下方插入。\n点击“取消”终止操作。",
                "选择插入方向",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question
            );

            if (direction == DialogResult.Cancel)
                return;

            Word.Table table = sel.Tables[1];
            Word.Row refRow;
            if (sel.Rows.Count > 0)
                refRow = sel.Rows[1];
            else
            {
                int rowIdx = sel.Information[Word.WdInformation.wdStartOfRangeRowNumber];
                refRow = table.Rows[rowIdx];
            }

            try
            {
                for (int i = 0; i < n; i++)
                {
                    if (direction == DialogResult.Yes)
                        refRow.Range.Rows.Add(refRow);        // 上方
                    else
                        refRow.Range.Rows.Add(refRow.Next);   // 下方
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("插入失败：" + ex.Message, "错误");
            }
        }


        private void 插入N列_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;

            if (sel == null || sel.Tables.Count == 0)
            {
                MessageBox.Show("请将光标放在表格内！", "提示");
                return;
            }

            // 输入要插入的列数
            string input = Interaction.InputBox("请输入要插入的列数：", "插入列", "1");
            if (string.IsNullOrWhiteSpace(input))
                return;

            if (!int.TryParse(input, out int n) || n <= 0)
            {
                MessageBox.Show("请输入有效的正整数！", "提示");
                return;
            }

            // 选择插入方向
            var direction = MessageBox.Show(
                "点击“是”在左侧插入，点击“否”在右侧插入。\n点击“取消”终止操作。",
                "选择插入方向",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question
            );

            if (direction == DialogResult.Cancel)
                return;

            Word.Table table = sel.Tables[1];
            Word.Column refCol;
            if (sel.Columns.Count > 0)
                refCol = sel.Columns[1];
            else
            {
                int colIdx = sel.Information[Word.WdInformation.wdStartOfRangeColumnNumber];
                refCol = table.Columns[colIdx];
            }

            try
            {
                for (int i = 0; i < n; i++)
                {
                    if (direction == DialogResult.Yes)
                        refCol.Select(); // 先选中目标列
                    else
                        refCol.Select();

                    if (direction == DialogResult.Yes)
                        refCol.Select(); // 选中左侧目标列
                    else
                        refCol.Select();

                    // 插入列
                    if (direction == DialogResult.Yes)
                        table.Columns.Add(refCol);
                    else
                        table.Columns.Add(refCol.Next);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("插入失败：" + ex.Message, "错误");
            }
        }

        private void 域名高亮_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            if (doc == null)
            {
                System.Windows.Forms.MessageBox.Show("未检测到文档。");
                return;
            }

            foreach (Word.Field field in doc.Fields)
            {
                string code = field.Code.Text.Trim();
                Word.Range fieldResult = field.Result;
                string fieldText = fieldResult.Text;

                // 1. 标准交叉引用：REF和HYPERLINK
                if (code.StartsWith("REF", StringComparison.OrdinalIgnoreCase) ||
                    code.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
                {
                    // 根据内容判定类型
                    if (fieldText.Contains("图"))
                    {
                        // 图，蓝色
                        fieldResult.Font.Color = Word.WdColor.wdColorBlue;
                    }
                    else if (fieldText.Contains("表"))
                    {
                        // 表，绿色
                        fieldResult.Font.Color = Word.WdColor.wdColorGreen;
                    }
                    else if (fieldText.Contains("公式"))
                    {
                        // 公式，红色
                        fieldResult.Font.Color = Word.WdColor.wdColorRed;
                    }
                    else
                    {
                        // 其它，紫色
                        fieldResult.Font.Color = Word.WdColor.wdColorBrown;
                    }
                }
                // 2. EndNote 文献引用（ADDIN类型，包含EN.CITE或EN.CITATION标记）
                else if (field.Type == Word.WdFieldType.wdFieldAddin &&
                         (code.Contains("EN.CITE") || code.Contains("EN.CITATION")))
                {
                    // 文献引用，高亮为金黄色
                    fieldResult.Font.Color = Word.WdColor.wdColorGold;
                }
            }

            System.Windows.Forms.MessageBox.Show("交叉引用与文献引用已高亮！");
        }

        private void 取消高亮_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            if (doc == null)
            {
                System.Windows.Forms.MessageBox.Show("未检测到文档。");
                return;
            }

            foreach (Word.Field field in doc.Fields)
            {
                string code = field.Code.Text.Trim();
                Word.Range fieldResult = field.Result;
                string fieldText = fieldResult.Text;

                // 1. 标准交叉引用：REF和HYPERLINK
                if (code.StartsWith("REF", StringComparison.OrdinalIgnoreCase) ||
                    code.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
                {
                    fieldResult.Font.Color = Word.WdColor.wdColorBlack;
                }
                // 2. EndNote 文献引用（ADDIN类型，包含EN.CITE或EN.CITATION标记）
                else if (field.Type == Word.WdFieldType.wdFieldAddin &&
                         (code.Contains("EN.CITE") || code.Contains("EN.CITATION")))
                {
                    // 文献引用，高亮为金黄色
                    fieldResult.Font.Color = Word.WdColor.wdColorBlack;
                }
            }

            System.Windows.Forms.MessageBox.Show("交叉引用与文献引用已取消高亮！");
        }

        private void 版本_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/eyinwei/WordMan_VSTO");
        }

        enum PictureNumberStyle
        {
            Arabic,     // 图 1
            Dash,       // 图 1-1
            Dot         // 图 1.1
        }
        PictureNumberStyle CurrentPicStyle = PictureNumberStyle.Arabic;
        private void 图注样式1_Click(object sender, RibbonControlEventArgs e)
        {
            图注样式1.Checked = true;
            图注样式2.Checked = false;
            图注样式3.Checked = false;
            CurrentPicStyle = PictureNumberStyle.Arabic;
        }
        private void 图注样式2_Click(object sender, RibbonControlEventArgs e)
        {
            图注样式1.Checked = false;
            图注样式2.Checked = true;
            图注样式3.Checked = false;
            CurrentPicStyle = PictureNumberStyle.Dash;
        }
        private void 图注样式3_Click(object sender, RibbonControlEventArgs e)
        {
            图注样式1.Checked = false;
            图注样式2.Checked = false;
            图注样式3.Checked = true;
            CurrentPicStyle = PictureNumberStyle.Dot;
        }

        private void 图片编号_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            var doc = app.ActiveDocument;

            HashSet<int> handledParagraphs = new HashSet<int>();
            List<Word.Paragraph> targetParas = new List<Word.Paragraph>();

            // 选区有图片
            foreach (Word.InlineShape ils in sel.Range.InlineShapes)
            {
                var para = ils.Range.Paragraphs.First;
                if (!handledParagraphs.Contains(para.Range.Start))
                {
                    targetParas.Add(para);
                    handledParagraphs.Add(para.Range.Start);
                }
            }
            foreach (Word.Shape s in sel.Range.ShapeRange)
            {
                var para = s.Anchor.Paragraphs.First;
                if (!handledParagraphs.Contains(para.Range.Start))
                {
                    targetParas.Add(para);
                    handledParagraphs.Add(para.Range.Start);
                }
            }

            // 若未选中图片，则取光标所在段落
            if (targetParas.Count == 0 && sel.Paragraphs.Count > 0)
            {
                var para = sel.Paragraphs.First;
                if (!handledParagraphs.Contains(para.Range.Start))
                {
                    targetParas.Add(para);
                    handledParagraphs.Add(para.Range.Start);
                }
            }



            // 必须逆序处理，防止段落因插入而错位
            for (int i = targetParas.Count - 1; i >= 0; i--)
            {
                InsertCaptionIfNotExists(targetParas[i], CurrentPicStyle);
            }
        }

        private void InsertCaptionIfNotExists(Word.Paragraph picPara, PictureNumberStyle numberStyle)
        {
            if (picPara == null) return;

            var doc = picPara.Range.Application.ActiveDocument;

            // 1. 检查后面是否已有题注
            var nextPara = picPara.Next() as Word.Paragraph;
            if (nextPara != null)
            {
                string nextText = nextPara.Range.Text.TrimStart();
                if ((nextPara.get_Style() is Word.Style style && style.NameLocal == "题注")
                    || nextText.StartsWith("图"))
                {
                    return; // 已有题注
                }
            }

            // 2. 插入空段并获得新段落
            var afterPicRange = picPara.Range.Duplicate;
            afterPicRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            afterPicRange.InsertParagraphAfter();
            var captionPara = picPara.Next() as Word.Paragraph;
            if (captionPara == null) return;

            // 3. 清空新段内容
            var captionRange = captionPara.Range.Duplicate;
            captionRange.End -= 1; // 去除段落标记
            captionRange.Text = "";

            // 4. 插入“图 ”（带空格）
            var insertRange = doc.Range(captionRange.Start, captionRange.Start);
            insertRange.InsertAfter("图 ");
            insertRange.SetRange(insertRange.Start + 2, insertRange.Start + 2); // 定位到空格后

            // 5. 插入编号
            switch (numberStyle)
            {
                case PictureNumberStyle.Arabic:
                    insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldSequence, "图 \\* ARABIC", false);
                    break;

                case PictureNumberStyle.Dash:
                case PictureNumberStyle.Dot:
                    {
                        // 插入章节号域
                        var styleRefField = insertRange.Fields.Add(
                            insertRange, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
                        // 跳出域
                        styleRefField.Result.Select();
                        var selection = insertRange.Application.Selection;
                        selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

                        // 插入分隔符
                        selection.TypeText(numberStyle == PictureNumberStyle.Dash ? "-" : ".");

                        selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

                        // 插入图序号
                        selection.Range.Fields.Add(
                            selection.Range, Word.WdFieldType.wdFieldSequence, "图 \\s 1", false);
                    }
                    break;
            }

            // 6. 设置样式为“题注”
            captionPara.set_Style("题注");
        }

        enum TableNumberStyle
        {
            Arabic,     // 表 1
            Dash,       // 表 1-1
            Dot         // 表 1.1
        }
        TableNumberStyle CurrentTableStyle = TableNumberStyle.Arabic;

        private void 表注样式1_Click(object sender, RibbonControlEventArgs e)
        {
            表注样式1.Checked = true;
            表注样式2.Checked = false;
            表注样式3.Checked = false;
            CurrentTableStyle = TableNumberStyle.Arabic;
        }
        private void 表注样式2_Click(object sender, RibbonControlEventArgs e)
        {
            表注样式1.Checked = false;
            表注样式2.Checked = true;
            表注样式3.Checked = false;
            CurrentTableStyle = TableNumberStyle.Dash;
        }
        private void 表注样式3_Click(object sender, RibbonControlEventArgs e)
        {
            表注样式1.Checked = false;
            表注样式2.Checked = false;
            表注样式3.Checked = true;
            CurrentTableStyle = TableNumberStyle.Dot;
        }
        private void 表格编号_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            var doc = app.ActiveDocument;

            HashSet<int> handledTables = new HashSet<int>();
            List<Word.Table> targetTables = new List<Word.Table>();

            // 1. 选区有表格则全部处理（修复多表格选择问题）
            // 关键修改：使用table.Range.InRange(sel.Range)判断表格是否完全在选区内
            foreach (Word.Table table in doc.Tables)
            {
                try
                {
                    // 检查表格是否被选中（表格范围在选区范围内）
                    if (table.Range.InRange(sel.Range) && !handledTables.Contains(table.Range.Start))
                    {
                        targetTables.Add(table);
                        handledTables.Add(table.Range.Start);
                    }
                }
                catch { } // 处理表格范围判断可能出现的异常
            }

            // 2. 若未选中表格，则处理光标所在表格
            if (targetTables.Count == 0 && sel.Tables.Count > 0)
            {
                var table = sel.Tables[1];
                if (!handledTables.Contains(table.Range.Start))
                {
                    targetTables.Add(table);
                    handledTables.Add(table.Range.Start);
                }
            }

            // 必须逆序处理，防止插入错位
            for (int i = targetTables.Count - 1; i >= 0; i--)
            {
                InsertTableCaptionIfNotExists(targetTables[i], CurrentTableStyle);
            }
        }
        private void InsertTableCaptionIfNotExists(Word.Table table, TableNumberStyle numberStyle)
        {
            if (table == null) return;
            var doc = table.Application.ActiveDocument;
            var app = doc.Application;

            // 获取表格真正的外部起始位置
            Word.Range tableRange = table.Range;
            int tableStart = tableRange.Start;

            // 检查是否在表格内部（调整到表格外部）
            bool isInFirstCell = table.Cell(1, 1).Range.InRange(tableRange);
            if (isInFirstCell)
            {
                tableStart = Math.Max(0, tableStart - 1);
            }

            // 1. 检查表格前是否已有题注
            Word.Paragraph prevPara = null;
            Word.Range beforeTableRange = doc.Range(0, tableStart);
            if (beforeTableRange.Paragraphs.Count > 0)
            {
                prevPara = beforeTableRange.Paragraphs[beforeTableRange.Paragraphs.Count];
                string prevText = prevPara.Range.Text.TrimStart();
                if ((prevPara.get_Style() is Word.Style style && style.NameLocal == "题注")
                    || prevText.StartsWith("表"))
                {
                    return; // 已有题注
                }
            }

            // 2. 保存原始表格位置用于定位
            int originalTablePosition = tableRange.Start;

            // 3. 插入题注段落（确保在表格外）
            // 关键修改：先清除可能的空内容，避免多余空行
            Word.Range insertRange = doc.Range(tableStart, tableStart);
            insertRange.Text = ""; // 清除插入位置可能的空内容
            insertRange.InsertParagraphBefore();

            // 4. 查找刚插入的题注段落
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

            // 5. 精确控制题注段落内容（解决空行问题）
            Word.Range captionRange = captionPara.Range.Duplicate;
            captionRange.End = captionRange.Start + 1; // 仅保留段落起始位置
            captionRange.Text = ""; // 彻底清空，避免默认空字符

            // 6. 插入"表 "和编号
            var fieldRange = doc.Range(captionRange.Start, captionRange.Start);
            fieldRange.InsertAfter("表 ");
            fieldRange.SetRange(fieldRange.Start + 2, fieldRange.Start + 2);

            switch (numberStyle)
            {
                case TableNumberStyle.Arabic:
                    fieldRange.Fields.Add(fieldRange, Word.WdFieldType.wdFieldSequence, "表 \\* ARABIC", false);
                    break;
                case TableNumberStyle.Dash:
                case TableNumberStyle.Dot:
                    var styleRefField = fieldRange.Fields.Add(
                        fieldRange, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
                    styleRefField.Result.Select();
                    var selection = fieldRange.Application.Selection;
                    selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

                    selection.TypeText(numberStyle == TableNumberStyle.Dash ? "-" : ".");

                    selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

                    selection.Range.Fields.Add(
                        selection.Range, Word.WdFieldType.wdFieldSequence, "表 \\s 1", false);
                    break;
            }

            // 7. 设置样式为题注并移除可能的多余空行
            captionPara.set_Style("题注");
            captionPara.SpaceAfter = 0; // 去除段后间距
            captionPara.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle; // 单倍行距
        }


        // 宽度刷相关字段
        private float copiedWidth = 0f;
        private Timer widthBrushTimer;
        private object lastSelectionStart = null;
        private object lastSelectionEnd = null;
        private DateTime lastActivityTime = DateTime.Now;
        private int lastSelectionHash = 0;

        private void 宽度刷_Click(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (toggleButton.Checked)
                    ActivateWidthBrush(app, toggleButton);
                else
                    DeactivateWidthBrush(app);
            }
            catch
            {
                toggleButton.Checked = false;
                DeactivateWidthBrush(Globals.ThisAddIn.Application);
            }
        }

        private void ActivateWidthBrush(Word.Application app, Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton)
        {
            var selection = app.Selection;
            if (selection.InlineShapes.Count == 0 && selection.ShapeRange.Count == 0)
            {
                toggleButton.Checked = false;
                return;
            }

            float width = 0f;
            if (selection.InlineShapes.Count > 0)
                width = selection.InlineShapes[1].Width;
            else if (selection.ShapeRange.Count > 0)
                width = selection.ShapeRange[1].Width;

            if (width <= 0)
            {
                toggleButton.Checked = false;
                return;
            }

            copiedWidth = width;
            lastActivityTime = DateTime.Now;
            lastSelectionHash = GetSelectionHash(selection);
            StartWidthBrushMonitoring(app);
            app.StatusBar = $"宽度刷已激活 - 宽度: {width}磅 - 按ESC或点击空白区域退出";
        }

        private void DeactivateWidthBrush(Word.Application app)
        {
            widthBrushTimer?.Stop();
            widthBrushTimer?.Dispose();
            widthBrushTimer = null;
            lastSelectionStart = null;
            lastSelectionEnd = null;
            lastSelectionHash = 0;
            宽度刷.Checked = false;
            try { app.StatusBar = ""; } catch { }
        }

        private void StartWidthBrushMonitoring(Word.Application app)
        {
            widthBrushTimer = new Timer();
            widthBrushTimer.Interval = 100;
            widthBrushTimer.Tick += (s, timerE) =>
            {
                if (!宽度刷.Checked) { DeactivateWidthBrush(app); return; }
                try
                {
                    var currentSelection = app.Selection;
                    int currentHash = GetSelectionHash(currentSelection);

                    // 检查退出条件
                    if (currentHash != lastSelectionHash)
                    {
                        bool isEmptyNow = currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                            (currentSelection.Type == Word.WdSelectionType.wdSelectionIP || currentSelection.Type == Word.WdSelectionType.wdSelectionNormal) &&
                            string.IsNullOrWhiteSpace(currentSelection.Text?.Replace("\r", ""));

                        TimeSpan timeSinceLastActivity = DateTime.Now - lastActivityTime;
                        if (isEmptyNow && timeSinceLastActivity.TotalMilliseconds < 500) { DeactivateWidthBrush(app); return; }

                        if (currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                            currentSelection.Paragraphs.Count > 0 &&
                            string.IsNullOrWhiteSpace(currentSelection.Paragraphs[1].Range.Text.Replace("\r", "")))
                        { DeactivateWidthBrush(app); return; }
                    }

                    // 长时间停留在空白区域退出
                    if (currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                        string.IsNullOrWhiteSpace(currentSelection.Text?.Replace("\r", "")) &&
                        (DateTime.Now - lastActivityTime).TotalMilliseconds > 1000)
                    { DeactivateWidthBrush(app); return; }

                    // 应用宽度
                    bool hasImageSelection = currentSelection.InlineShapes.Count > 0 || currentSelection.ShapeRange.Count > 0;
                    if (hasImageSelection)
                    {
                        bool isNewSelection = false;
                        try
                        {
                            if (lastSelectionStart == null || lastSelectionEnd == null ||
                                !lastSelectionStart.Equals(currentSelection.Start) || !lastSelectionEnd.Equals(currentSelection.End))
                            {
                                isNewSelection = true;
                                lastSelectionStart = currentSelection.Start;
                                lastSelectionEnd = currentSelection.End;
                                lastActivityTime = DateTime.Now;
                                lastSelectionHash = currentHash;
                            }
                        }
                        catch { isNewSelection = true; lastActivityTime = DateTime.Now; lastSelectionHash = currentHash; }

                        if (isNewSelection && ApplyWidthToSelection(currentSelection))
                            app.StatusBar = $"宽度刷: 已应用宽度 {copiedWidth}磅 - 按ESC或点击空白区域退出";
                    }
                    else if (currentHash != lastSelectionHash)
                    {
                        lastActivityTime = DateTime.Now;
                        lastSelectionHash = currentHash;
                    }
                }
                catch { }
            };
            widthBrushTimer.Start();
        }

        private int GetSelectionHash(Word.Selection selection)
        {
            try
            {
                int hash = 0;
                hash ^= selection.Start.GetHashCode();
                hash ^= selection.End.GetHashCode();
                hash ^= selection.Type.GetHashCode();
                hash ^= selection.InlineShapes.Count.GetHashCode();
                hash ^= selection.ShapeRange.Count.GetHashCode();
                if (!string.IsNullOrEmpty(selection.Text))
                    hash ^= selection.Text.GetHashCode();
                return hash;
            }
            catch { return DateTime.Now.Millisecond; }
        }

        private bool ApplyWidthToSelection(Word.Selection selection)
        {
            bool applied = false;
            try
            {
                if (selection.InlineShapes.Count > 0)
                {
                    foreach (Word.InlineShape shape in selection.InlineShapes)
                    {
                        if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeChart ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeSmartArt)
                        {
                            shape.Width = copiedWidth;
                            applied = true;
                        }
                    }
                }

                if (selection.ShapeRange.Count > 0)
                {
                    foreach (Word.Shape shape in selection.ShapeRange)
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoChart ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            shape.Width = copiedWidth;
                            applied = true;
                        }
                    }
                }
            }
            catch { }
            return applied;
        }


        // 高度刷相关字段
        private float copiedHeight = 0f;
        private Timer heightBrushTimer;
        private object lastHeightSelectionStart = null;
        private object lastHeightSelectionEnd = null;
        private DateTime lastHeightActivityTime = DateTime.Now;
        private int lastHeightSelectionHash = 0;

        private void 高度刷_Click(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (toggleButton.Checked)
                    ActivateHeightBrush(app, toggleButton);
                else
                    DeactivateHeightBrush(app);
            }
            catch
            {
                toggleButton.Checked = false;
                DeactivateHeightBrush(Globals.ThisAddIn.Application);
            }
        }

        private void ActivateHeightBrush(Word.Application app, Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton)
        {
            var selection = app.Selection;
            if (selection.InlineShapes.Count == 0 && selection.ShapeRange.Count == 0)
            {
                toggleButton.Checked = false;
                return;
            }

            float height = 0f;
            if (selection.InlineShapes.Count > 0)
                height = selection.InlineShapes[1].Height;
            else if (selection.ShapeRange.Count > 0)
                height = selection.ShapeRange[1].Height;

            if (height <= 0)
            {
                toggleButton.Checked = false;
                return;
            }

            copiedHeight = height;
            lastHeightActivityTime = DateTime.Now;
            lastHeightSelectionHash = GetHeightSelectionHash(selection);
            StartHeightBrushMonitoring(app);
            app.StatusBar = $"高度刷已激活 - 高度: {height}磅 - 按ESC或点击空白区域退出";
        }

        private void DeactivateHeightBrush(Word.Application app)
        {
            heightBrushTimer?.Stop();
            heightBrushTimer?.Dispose();
            heightBrushTimer = null;
            lastHeightSelectionStart = null;
            lastHeightSelectionEnd = null;
            lastHeightSelectionHash = 0;
            高度刷.Checked = false;
            try { app.StatusBar = ""; } catch { }
        }

        private void StartHeightBrushMonitoring(Word.Application app)
        {
            heightBrushTimer = new Timer();
            heightBrushTimer.Interval = 100;
            heightBrushTimer.Tick += (s, timerE) =>
            {
                if (!高度刷.Checked) { DeactivateHeightBrush(app); return; }
                try
                {
                    var currentSelection = app.Selection;
                    int currentHash = GetHeightSelectionHash(currentSelection);

                    // 检查退出条件
                    if (currentHash != lastHeightSelectionHash)
                    {
                        bool isEmptyNow = currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                            (currentSelection.Type == Word.WdSelectionType.wdSelectionIP || currentSelection.Type == Word.WdSelectionType.wdSelectionNormal) &&
                            string.IsNullOrWhiteSpace(currentSelection.Text?.Replace("\r", ""));

                        TimeSpan timeSinceLastActivity = DateTime.Now - lastHeightActivityTime;
                        if (isEmptyNow && timeSinceLastActivity.TotalMilliseconds < 500) { DeactivateHeightBrush(app); return; }

                        if (currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                            currentSelection.Paragraphs.Count > 0 &&
                            string.IsNullOrWhiteSpace(currentSelection.Paragraphs[1].Range.Text.Replace("\r", "")))
                        { DeactivateHeightBrush(app); return; }
                    }

                    // 长时间停留在空白区域退出
                    if (currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                        string.IsNullOrWhiteSpace(currentSelection.Text?.Replace("\r", "")) &&
                        (DateTime.Now - lastHeightActivityTime).TotalMilliseconds > 1000)
                    { DeactivateHeightBrush(app); return; }

                    // 应用高度
                    bool hasImageSelection = currentSelection.InlineShapes.Count > 0 || currentSelection.ShapeRange.Count > 0;
                    if (hasImageSelection)
                    {
                        bool isNewSelection = false;
                        try
                        {
                            if (lastHeightSelectionStart == null || lastHeightSelectionEnd == null ||
                                !lastHeightSelectionStart.Equals(currentSelection.Start) || !lastHeightSelectionEnd.Equals(currentSelection.End))
                            {
                                isNewSelection = true;
                                lastHeightSelectionStart = currentSelection.Start;
                                lastHeightSelectionEnd = currentSelection.End;
                                lastHeightActivityTime = DateTime.Now;
                                lastHeightSelectionHash = currentHash;
                            }
                        }
                        catch { isNewSelection = true; lastHeightActivityTime = DateTime.Now; lastHeightSelectionHash = currentHash; }

                        if (isNewSelection && ApplyHeightToSelection(currentSelection))
                            app.StatusBar = $"高度刷: 已应用高度 {copiedHeight}磅 - 按ESC或点击空白区域退出";
                    }
                    else if (currentHash != lastHeightSelectionHash)
                    {
                        lastHeightActivityTime = DateTime.Now;
                        lastHeightSelectionHash = currentHash;
                    }
                }
                catch { }
            };
            heightBrushTimer.Start();
        }

        private int GetHeightSelectionHash(Word.Selection selection)
        {
            try
            {
                int hash = 0;
                hash ^= selection.Start.GetHashCode();
                hash ^= selection.End.GetHashCode();
                hash ^= selection.Type.GetHashCode();
                hash ^= selection.InlineShapes.Count.GetHashCode();
                hash ^= selection.ShapeRange.Count.GetHashCode();
                if (!string.IsNullOrEmpty(selection.Text))
                    hash ^= selection.Text.GetHashCode();
                return hash;
            }
            catch { return DateTime.Now.Millisecond; }
        }

        private bool ApplyHeightToSelection(Word.Selection selection)
        {
            bool applied = false;
            try
            {
                if (selection.InlineShapes.Count > 0)
                {
                    foreach (Word.InlineShape shape in selection.InlineShapes)
                    {
                        if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeChart ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeSmartArt)
                        {
                            shape.Height = copiedHeight;
                            applied = true;
                        }
                    }
                }

                if (selection.ShapeRange.Count > 0)
                {
                    foreach (Word.Shape shape in selection.ShapeRange)
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoChart ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            shape.Height = copiedHeight;
                            applied = true;
                        }
                    }
                }
            }
            catch { }
            return applied;
        }

        // 在Cleanup方法中添加高度刷的清理
        public void Cleanup()
        {
            widthBrushTimer?.Stop();
            widthBrushTimer?.Dispose();
            widthBrushTimer = null;

            heightBrushTimer?.Stop();
            heightBrushTimer?.Dispose();
            heightBrushTimer = null;
        }




        // 排版按钮点击事件
        private void TypesettingButton_Click(object sender, RibbonControlEventArgs e)
        {
            // 仅一行：调用任务窗格的静态方法，剩下的全由任务窗格自己处理
            TypesettingTaskPane.TriggerShowOrHide();
        }

        // 位图化按钮点击事件
        private void 位图化_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.Selection;
                
                // 检查是否选中了单个图形并转换
                if (selection.InlineShapes.Count == 1)
                    ConvertToBitmap(selection.InlineShapes[1], app);
                else if (selection.ShapeRange.Count == 1)
                    ConvertToBitmap(selection.ShapeRange[1], app);
            }
            catch
            {
                // 静默处理错误，不显示提示
            }
        }

        // 转换方法
        private void ConvertToBitmap(object shape, Word.Application app)
        {
            // 如果不是位图则转换
            if (!IsBitmap(shape))
            {
                // 统一处理：选择、复制、删除、粘贴
                if (shape is Word.InlineShape inlineShape)
                {
                    inlineShape.Select();
                    app.Selection.Copy();
                    inlineShape.Delete();
                }
                else if (shape is Word.Shape wordShape)
                {
                    wordShape.Select();
                    app.Selection.Copy();
                    wordShape.Delete();
                }
                app.Selection.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteBitmap);
            }
        }

        // 判断是否为位图
        private bool IsBitmap(object shape)
        {
            if (shape is Word.InlineShape inlineShape)
            {
                var type = inlineShape.Type;
                return (type == Word.WdInlineShapeType.wdInlineShapePicture ||
                        type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture) &&
                       inlineShape.PictureFormat != null;
            }
            else if (shape is Word.Shape wordShape)
            {
                var type = wordShape.Type;
                return (type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                        type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture) &&
                       wordShape.PictureFormat != null;
            }
            return false;
        }



        // 快速密级相关功能
        private void 公开_Click(object sender, RibbonControlEventArgs e)
        {
            AddSecurityLevel("公开");
        }

        private void 内部_Click(object sender, RibbonControlEventArgs e)
        {
            AddSecurityLevel("内部★");
        }

        private void 移除密级_Click(object sender, RibbonControlEventArgs e)
        {
            RemoveSecurityLevelFromCurrentPage();
        }

        // 添加密级标签
        private void AddSecurityLevel(string levelText)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                var selection = app.Selection;

                // 先移除当前页的密级标签
                RemoveSecurityLevelFromCurrentPage();

                // 获取当前页信息
                int currentPage = selection.Information[Word.WdInformation.wdActiveEndPageNumber];
                
                // 获取页面设置信息
                var pageSetup = doc.PageSetup;
                float leftMargin = pageSetup.LeftMargin;
                float topMargin = pageSetup.TopMargin;
                
                // 在页边距外侧添加密级标签
                // 移动到当前页开始位置
                selection.HomeKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);
                
                // 使用Shapes.AddTextbox方法创建文本框，锚点到当前选区
                var textBox = doc.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 20);
                
                // 设置文本框内容
                textBox.TextFrame.TextRange.Text = levelText;
                
                // 设置文本框格式
                var textRange = textBox.TextFrame.TextRange;
                textRange.Font.Name = "黑体";
                textRange.Font.Size = 12; // 小三号字体
                textRange.Font.Bold = 1;
                textRange.Font.Color = Word.WdColor.wdColorBlack; // 黑色字体
                
                // 设置文本框边框和背景
                textBox.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // 无边框
                textBox.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse; // 无背景
                // 先设置文本框大小
                textBox.Width = app.CentimetersToPoints(3.0f);  // 3厘米宽
                textBox.Height = app.CentimetersToPoints(0.8f); // 0.8厘米高
                
                // 设置文本框位置
                // 水平方向：相对于页边距左对齐
                textBox.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                textBox.Left = 0; // 页边距起始位置
                
                // 垂直方向：相对于页边距，文本框底部与页边距对齐
                textBox.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                textBox.Top = -textBox.Height; // 上页边距位置减去文本框高度
                
                textBox.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加密级标签失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 移除当前页密级标签
        private void RemoveSecurityLevelFromCurrentPage()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                var selection = app.Selection;
                
                // 获取当前页信息
                int currentPage = selection.Information[Word.WdInformation.wdActiveEndPageNumber];
                
                // 查找并删除当前页包含密级文本的文本框
                foreach (Word.Shape shape in doc.Shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                    {
                        string text = shape.TextFrame.TextRange.Text.Trim();
                        if (text == "公开" || text == "内部★" || text.Contains("密级"))
                        {
                            // 检查文本框是否在当前页
                            try
                            {
                                int shapePage = shape.Anchor.Information[Word.WdInformation.wdActiveEndPageNumber];
                                if (shapePage == currentPage)
                                {
                                    shape.Delete();
                                }
                            }
                            catch
                            {
                                // 如果无法确定页数，也删除（可能是浮动文本框）
                                shape.Delete();
                            }
                        }
                    }
                }
            }
            catch
            {
                // 静默处理错误，避免影响用户体验
            }
        }

        // 移除所有密级标签（保留原方法用于其他用途）
        private void RemoveSecurityLevel()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                
                // 查找并删除所有包含密级文本的文本框
                foreach (Word.Shape shape in doc.Shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                    {
                        string text = shape.TextFrame.TextRange.Text.Trim();
                        if (text == "公开" || text == "内部★" || text.Contains("密级"))
                        {
                            shape.Delete();
                        }
                    }
                }
            }
            catch
            {
                // 静默处理错误，避免影响用户体验
            }
        }

        // 文档样式设置按钮点击事件
        private void 样式设置_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var styleSettings = new StyleSettings();
                styleSettings.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开样式设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 文档拆分按钮点击事件
        private void 文档拆分_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var splitter = new DocumentSplitter(Globals.ThisAddIn.Application);
                splitter.ShowSplitDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文档拆分失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 多级列表按钮点击事件
        private void 多级列表_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var multiLevelForm = new MultiLevelListForm();
                multiLevelForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开多级列表设置失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 文档合并按钮点击事件
        private void 文档合并_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var merger = new DocumentMerger((Microsoft.Office.Interop.Word.Application)Globals.ThisAddIn.Application);
                merger.ShowMergeDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文档合并失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}







