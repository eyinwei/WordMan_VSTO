using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan
{
    public class TableProcessor
    {
        /// <summary>
        /// 创建三线表
        /// </summary>
        public void CreateThreeLineTable()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;

            // 1. 创建3x2表格
            Word.Table table = sel.Tables.Add(sel.Range, 3, 3);

            // 2. 选中整个表格
            table.Select();

            // 3. 调用已有的设为三线方法
            SetThreeLineTable(table);
        }

        /// <summary>
        /// 将表格设为三线表格式
        /// </summary>
        /// <param name="table">要设置的表格</param>
        public void SetThreeLineTable(Word.Table table)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            
            if (table == null)
                return;

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

                    // 第一行：加下边细线（即三线表"中线"）
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

        /// <summary>
        /// 插入N行
        /// </summary>
        public void InsertNRows()
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

        /// <summary>
        /// 插入N列
        /// </summary>
        public void InsertNColumns()
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
    }
}