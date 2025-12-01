using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
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
            try
            {
                var sel = Globals.ThisAddIn.Application.Selection;
                if (sel == null || sel.Range == null)
                {
                    MessageBox.Show("无法创建表格，请检查文档状态。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                Word.Table table = sel.Tables.Add(sel.Range, 3, 3);
                table.Select();
                SetThreeLineTable(table);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建三线表失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 将当前选中的表格设为三线表格式
        /// </summary>
        public void SetCurrentTableToThreeLine()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            if (sel == null || sel.Tables.Count == 0)
                return;

            Word.Table table = sel.Tables[1];
            SetThreeLineTable(table);
        }

        /// <summary>
        /// 将表格设为三线表格式
        /// </summary>
        /// <param name="table">要设置的表格</param>
        public void SetThreeLineTable(Word.Table table)
        {
            if (table == null)
                return;

            // 找出最小和最大行号（因为有合并单元格，不能用Rows.Count）
            int firstRowIndex = int.MaxValue;
            int lastRowIndex = int.MinValue;
            foreach (Word.Cell cell in table.Range.Cells)
            {
                if (cell.RowIndex < firstRowIndex)
                    firstRowIndex = cell.RowIndex;
                if (cell.RowIndex > lastRowIndex)
                    lastRowIndex = cell.RowIndex;
            }

            // 合并遍历：清除所有边框并设置三线
            Word.WdBorderType[] borderTypes = new[]
            {
                Word.WdBorderType.wdBorderLeft,
                Word.WdBorderType.wdBorderRight,
                Word.WdBorderType.wdBorderTop,
                Word.WdBorderType.wdBorderBottom
            };

            foreach (Word.Cell cell in table.Range.Cells)
            {
                // 清除所有边框
                foreach (Word.WdBorderType borderType in borderTypes)
                {
                    cell.Borders[borderType].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                }

                // 设置三线表边框
                if (cell.RowIndex == firstRowIndex)
                {
                    SetBorder(cell, Word.WdBorderType.wdBorderTop, Word.WdLineWidth.wdLineWidth150pt);
                    SetBorder(cell, Word.WdBorderType.wdBorderBottom, Word.WdLineWidth.wdLineWidth075pt);
                }
                if (cell.RowIndex == lastRowIndex)
                {
                    SetBorder(cell, Word.WdBorderType.wdBorderBottom, Word.WdLineWidth.wdLineWidth150pt);
                }
            }

            // 设置表格格式
            ApplyTableFormatting(table);
        }

        private void SetBorder(Word.Cell cell, Word.WdBorderType borderType, Word.WdLineWidth lineWidth)
        {
            cell.Borders[borderType].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
            cell.Borders[borderType].LineWidth = lineWidth;
        }

        private void ApplyTableFormatting(Word.Table table)
        {
            table.Range.Font.Size = 10.5f;
            table.Range.Font.NameFarEast = "宋体";
            table.Range.Font.Name = "Times New Roman";
            table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100f;
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
        }

        /// <summary>
        /// 插入N行
        /// </summary>
        public void InsertNRows()
        {
            var sel = Globals.ThisAddIn.Application.Selection;
            if (!ValidateTableSelection(sel))
                return;

            int n = GetInsertCount("请输入要插入的行数：", "插入行");
            if (n <= 0)
                return;

            DialogResult direction = GetInsertDirection(
                "点击\"是\"在上方插入，点击\"否\"在下方插入。\n点击\"取消\"终止操作。",
                "选择插入方向");
            if (direction == DialogResult.Cancel)
                return;

            Word.Table table = sel.Tables[1];
            Word.Row refRow = GetReferenceRow(sel, table);

            try
            {
                for (int i = 0; i < n; i++)
                {
                    if (direction == DialogResult.Yes)
                        refRow.Range.Rows.Add(refRow);
                    else
                        refRow.Range.Rows.Add(refRow.Next);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("插入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 插入N列
        /// </summary>
        public void InsertNColumns()
        {
            var sel = Globals.ThisAddIn.Application.Selection;
            if (!ValidateTableSelection(sel))
                return;

            int n = GetInsertCount("请输入要插入的列数：", "插入列");
            if (n <= 0)
                return;

            DialogResult direction = GetInsertDirection(
                "点击\"是\"在左侧插入，点击\"否\"在右侧插入。\n点击\"取消\"终止操作。",
                "选择插入方向");
            if (direction == DialogResult.Cancel)
                return;

            Word.Table table = sel.Tables[1];
            Word.Column refCol = GetReferenceColumn(sel, table);

            try
            {
                for (int i = 0; i < n; i++)
                {
                    if (direction == DialogResult.Yes)
                        table.Columns.Add(refCol);
                    else
                        table.Columns.Add(refCol.Next);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("插入失败：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region 辅助方法
        private bool ValidateTableSelection(Word.Selection sel)
        {
            if (sel == null || sel.Tables.Count == 0)
            {
                MessageBox.Show("请将光标放在表格内！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            return true;
        }

        private int GetInsertCount(string prompt, string title)
        {
            string input = Interaction.InputBox(prompt, title, "1");
            if (string.IsNullOrWhiteSpace(input))
                return 0;

            if (!int.TryParse(input, out int n) || n <= 0)
            {
                MessageBox.Show("请输入有效的正整数！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            return n;
        }

        private DialogResult GetInsertDirection(string message, string title)
        {
            return MessageBox.Show(message, title, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        }

        private Word.Row GetReferenceRow(Word.Selection sel, Word.Table table)
        {
            if (sel.Rows.Count > 0)
                return sel.Rows[1];
            
            int rowIdx = sel.Information[Word.WdInformation.wdStartOfRangeRowNumber];
            return table.Rows[rowIdx];
        }

        private Word.Column GetReferenceColumn(Word.Selection sel, Word.Table table)
        {
            if (sel.Columns.Count > 0)
                return sel.Columns[1];
            
            int colIdx = sel.Information[Word.WdInformation.wdStartOfRangeColumnNumber];
            return table.Columns[colIdx];
        }
        #endregion

        /// <summary>
        /// 重复标题行 - 执行切换命令
        /// </summary>
        public void RepeatHeaderRows()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                
                if (sel == null || sel.Tables.Count == 0)
                {
                    MessageBox.Show("请将光标放在表格中。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // 直接执行Word内置命令，Word会自动管理按钮状态
                app.CommandBars.ExecuteMso("TableRepeatHeaderRows");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行重复标题行失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 获取当前表格的重复标题行状态
        /// </summary>
        public bool GetRepeatHeaderRowsState()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                
                if (sel == null || sel.Tables.Count == 0)
                    return false;
                
                Word.Table table = sel.Tables[1];
                
                // 检查第一行是否设置为标题行
                return table.Rows[1].HeadingFormat != 0;
            }
            catch
            {
                return false;
            }
        }
    }
}