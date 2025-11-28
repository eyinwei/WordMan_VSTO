using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace WordMan
{
    public class CaptionManager
    {
        #region 枚举定义
        public enum PictureNumberStyle
        {
            Arabic,     // 图 1
            Dash,       // 图 1-1
            Dot         // 图 1.1
        }

        public enum TableNumberStyle
        {
            Arabic,     // 表 1
            Dash,       // 表 1-1
            Dot         // 表 1.1
        }

        public enum FormulaNumberStyle
        {
            Parenthesis1,    // (1)
            Parenthesis1_1,  // (1-1)
            Parenthesis1_1dot// (1.1)
        }
        #endregion

        #region 公式编号相关方法
        public static Word.Table CreateFormulaTable(Word.Selection sel, Word.Application app)
        {
            // 插入一行三列的表格
            Word.Table table = sel.Tables.Add(sel.Range, 1, 3);

            // 设置表格属性为无边框
            table.Borders.Enable = 0;

            // 计算并设置列宽
            float pageWidth = sel.PageSetup.PageWidth - sel.PageSetup.LeftMargin - sel.PageSetup.RightMargin;
            float[] columnWidths = { pageWidth * 0.15f, pageWidth * 0.7f, pageWidth * 0.15f }; // 左15%, 中70%, 右15%

            for (int i = 0; i < 3; i++)
            {
                table.Columns[i + 1].Width = app.CentimetersToPoints(columnWidths[i] / 28.35f);
            }

            // 设置单元格对齐方式
            Word.WdParagraphAlignment[] alignments =
            {
                Word.WdParagraphAlignment.wdAlignParagraphLeft,   // 第一列：左对齐
                Word.WdParagraphAlignment.wdAlignParagraphCenter, // 第二列：居中
                Word.WdParagraphAlignment.wdAlignParagraphRight   // 第三列：右对齐
            };

            for (int i = 0; i < 3; i++)
            {
                table.Cell(1, i + 1).Range.ParagraphFormat.Alignment = alignments[i];
            }

            return table;
        }

        public static void InsertFormulaNumber(Word.Table table, Word.Selection sel, FormulaNumberStyle currentStyle)
        {
            const string leftBracket = "(";
            const string rightBracket = ")";
            const string seqName = "公式";

            // 移动到第三列
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

        #region 图片编号相关方法
        public static void InsertPictureCaption(Word.Paragraph picPara, PictureNumberStyle numberStyle)
        {
            if (picPara == null) return;

            var doc = picPara.Range.Application.ActiveDocument;

            // 1. 检查后面是否已有题注
            var nextPara = picPara.Next() as Word.Paragraph;
            if (nextPara != null)
            {
                string nextText = nextPara.Range.Text.Trim();
                // 检查是否已经有题注样式或以"图"开头的文本
                if (!string.IsNullOrEmpty(nextText))
                {
                    if ((nextPara.get_Style() is Word.Style style && style.NameLocal == "题注")
                        || nextText.StartsWith("图"))
                    {
                        return; // 已有题注
                    }
                }
            }

            // 2. 保存原始段落位置用于定位
            int originalPicPosition = picPara.Range.End;

            // 3. 插入空段并获得新段落
            var afterPicRange = picPara.Range.Duplicate;
            afterPicRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            afterPicRange.InsertParagraphAfter();

            // 4. 查找刚插入的题注段落
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

            // 5. 精确控制题注段落内容（解决空行问题）
            Word.Range captionRange = captionPara.Range.Duplicate;
            captionRange.End = captionRange.Start + 1; // 仅保留段落起始位置
            captionRange.Text = ""; // 彻底清空，避免默认空字符

            // 6. 插入"图 "（带空格）
            var insertRange = doc.Range(captionRange.Start, captionRange.Start);
            insertRange.InsertAfter("图 ");
            insertRange.SetRange(insertRange.Start + 2, insertRange.Start + 2); // 定位到空格后

            // 7. 插入编号
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

            // 8. 设置样式为"题注"
            captionPara.set_Style("题注");
        }
        #endregion

        #region 表格编号相关方法
        public static void InsertTableCaption(Word.Table table, TableNumberStyle numberStyle)
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

                    // 插入分隔符
                    selection.TypeText(numberStyle == TableNumberStyle.Dash ? "-" : ".");

                    selection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

                    // 插入表序号
                    selection.Range.Fields.Add(
                        selection.Range, Word.WdFieldType.wdFieldSequence, "表 \\s 1", false);
                    break;
            }

            // 7. 设置样式为"题注"
            captionPara.set_Style("题注");
        }
        #endregion

        #region 交叉引用相关方法
        public class CaptionInfo
        {
            public string Identifier { get; set; }  // 标签类型，如"图", "表"
            public string Number { get; set; }      // 编号，如"1", "1-1"
            public string FullText { get; set; }    // 完整题注文本
        }

        public static CaptionInfo FindCaptionInfo(Word.Range range)
        {
            // 获取当前段落
            Word.Paragraph paragraph = range.Paragraphs[1];
            Word.Range paraRange = paragraph.Range;

            // 在段落中查找域
            foreach (Word.Field field in paraRange.Fields)
            {
                if (field.Type == Word.WdFieldType.wdFieldSequence)
                {
                    // 获取SEQ域的标识符（如"图"、"表"等）
                    string fieldCode = field.Code.Text;
                    string identifier = ExtractIdentifierFromFieldCode(fieldCode);

                    // 获取题注的完整文本（显示文本）
                    string captionText = field.Result.Text.Trim();

                    // 获取题注编号（从显示文本中提取数字）
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

        private static string ExtractIdentifierFromFieldCode(string fieldCode)
        {
            // 解析域代码，提取标识符
            // 示例: " SEQ 图 \* ARABIC " -> "图"
            fieldCode = fieldCode.Trim();

            if (fieldCode.StartsWith("SEQ", StringComparison.OrdinalIgnoreCase))
            {
                string remaining = fieldCode.Substring(3).Trim();

                // 找到第一个空格或反斜杠
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

        private static string ExtractNumberFromCaption(string captionText)
        {
            // 从题注文本中提取编号
            // 例如："图 1-1" -> "1-1"
            if (string.IsNullOrEmpty(captionText))
                return string.Empty;

            // 移除标签前缀（如"图 "、"表 "等）
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

        public static void InsertCrossReferenceAtOriginalPosition(Word.Range originalRange, CaptionInfo captionInfo)
        {
            // 返回到原始位置并插入交叉引用
            Globals.ThisAddIn.Application.Selection.SetRange(
                originalRange.Start, originalRange.End);

            // 直接使用Word的交叉引用功能
            object referenceType = captionInfo.Identifier; // 如"图", "表"

            Word.WdReferenceKind referenceKind = (captionInfo.Identifier == "公式" || captionInfo.Identifier == "式" || captionInfo.Identifier == "EQ") ?
            Word.WdReferenceKind.wdEntireCaption : Word.WdReferenceKind.wdOnlyLabelAndNumber;

            object referenceItem = captionInfo.Number; // 编号

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
    }
}