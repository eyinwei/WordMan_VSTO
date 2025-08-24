using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.VisualBasic;    
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    public partial class MainRibbon
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
                    CreateBookmarks: Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateWordBookmarks,
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

            var en2cn = new Dictionary<string, string>
            {
                {";", "；"}, {":", "："}, {",", "，"}, {".", "。"}, {"?", "？"}, {"!", "！"},
                {"(", "（"}, {")", "）"}, {"[", "【"}, {"]", "】"}, {"<", "《"}, {">", "》"}
            };
            var cn2en = new Dictionary<string, string>
            {
                {"；", ";"}, {"：", ":"}, {"，", ","}, {"。", "."}, {"？", "?"}, {"！", "!"},
                {"（", "("}, {"）", ")"}, {"【", "["}, {"】", "]"}, {"《", "<"}, {"》", ">"},
                {"“", "\""}, {"”", "\""}, {"‘", "'"}, {"’", "'"}, {"　", " "},
                 // 补充常见全角符号
                {"／", "/"}, {"＂", "\""}, {"＇", "'"}, {"＆", "&"}, {"＃", "#"},
                {"％", "%"}, {"＊", "*"}, {"＋", "+"}, {"－", "-"}, {"＝", "="},
                {"＠", "@"}, {"＄", "$"}, {"＾", "^"}, {"＿", "_"}, {"｀", "`"},
                {"｜", "|"}, {"＼", "\\"}, {"～", "~"},
                {"µ", "μ"},    // 微符号（U+00B5）→ 希腊小写mu（U+03BC）
                {"Ω", "Ω"},    // 欧姆符号（U+2126）→ 希腊大写Omega（U+03A9）
                {"℧", "Ʊ"},    // 倒欧姆符号（U+2127）→ 拉丁大写Ʊ（近似），如需其他可自定义
                {"∑", "Σ"},    // 求和符号（U+2211）→ 希腊大写Sigma（U+03A3）
                {"∆", "Δ"},    // 增量符号（U+2206）→ 希腊大写Delta（U+0394）
                {"∏", "Π"}    // 连乘符号（U+220F）→ 希腊大写Pi（U+03A0）
            };

            var dict = englishToChinese ? en2cn : cn2en;

            foreach (var pair in dict)
            {
                rng.Find.ClearFormatting();
                rng.Find.Text = pair.Key;
                rng.Find.Replacement.ClearFormatting();
                rng.Find.Replacement.Text = pair.Value;
                rng.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            }

            // 只在英标转中标时做成对引号替换
            if (englishToChinese)
            {
                void ReplacePairQuotes(Range range, string from, string left, string right)
                {
                    bool isLeft = true;
                    Range search = range.Duplicate;
                    search.Find.ClearFormatting();
                    search.Find.Text = from;
                    search.Find.Forward = true;
                    search.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    while (search.Find.Execute())
                    {
                        Range hit = app.ActiveDocument.Range(search.Start, search.Start + from.Length);
                        hit.Text = isLeft ? left : right;
                        isLeft = !isLeft;
                        search.SetRange(hit.End, range.End);
                        search.Find.Text = from;
                    }
                }

                ReplacePairQuotes(rng, "\"", "“", "”");
                ReplacePairQuotes(rng, "'", "‘", "’");
            }
        }

        // 调用方式
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

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            toggleButton1.Checked = true;
            toggleButton2.Checked = false;
            toggleButton3.Checked = false;
            CurrentStyle = FormulaNumberStyle.Parenthesis1;
        }
        private void toggleButton2_Click(object sender, RibbonControlEventArgs e)
        {
            toggleButton1.Checked = false;
            toggleButton2.Checked = true;
            toggleButton3.Checked = false;
            CurrentStyle = FormulaNumberStyle.Parenthesis1_1;
        }
        private void toggleButton3_Click(object sender, RibbonControlEventArgs e)
        {
            toggleButton1.Checked = false;
            toggleButton2.Checked = false;
            toggleButton3.Checked = true;
            CurrentStyle = FormulaNumberStyle.Parenthesis1_1dot;
        }

        private void 编号_Click(object sender, RibbonControlEventArgs e)
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
                    insertRange = seqField.Result.Duplicate;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    break;

                case FormulaNumberStyle.Parenthesis1_1:
                    // 公式（1-1） ==> 第一个1由STYLEREF，第2个1由SEQ
                    var srField2 = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
                    insertRange = srField2.Result.Duplicate;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    insertRange.InsertAfter("-");
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    var seqField2 = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldSequence, seqName, false);
                    insertRange = seqField2.Result.Duplicate;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    break;

                case FormulaNumberStyle.Parenthesis1_1dot:
                    // 公式（1.1） ==> 第一个1由STYLEREF，第2个1由SEQ
                    var srField3 = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldStyleRef, "1 \\s", false);
                    insertRange = srField3.Result.Duplicate;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    insertRange.InsertAfter(".");
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                    var seqField3 = insertRange.Fields.Add(insertRange, Word.WdFieldType.wdFieldSequence, seqName, false);
                    insertRange = seqField3.Result.Duplicate;
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    break;
            }

            insertRange.InsertAfter(rightBracket);
        }




        private void 三线表_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
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

    }
}




