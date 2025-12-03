using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan
{
    public class TableProcessor
    {
        #region 创建表格方法
        public void CreateThreeLineTableStyle()
        {
            CreateTableWithStyle(4, 4, SetThreeLineTableStyle, "创建三线表失败");
        }

        public void CreateGBTableStyle()
        {
            CreateTableWithStyle(4, 4, SetGBTableStyle, "创建国标表格失败");
        }

        public void CreateNoBorderTableStyle()
        {
            CreateTableWithStyle(2, 2, SetNoBorderTableStyle, "创建无框线表格失败");
        }

        public void CreateTable()
        {
            CreateTableWithStyle(3, 3, null, "创建表格失败");
        }

        private void CreateTableWithStyle(int rows, int columns, Action<Word.Table> applyStyle, string errorMessage)
        {
            try
            {
                var sel = Globals.ThisAddIn.Application.Selection;
                if (sel == null || sel.Range == null)
                {
                    MessageBox.Show("无法创建表格，请检查文档状态。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                Word.Table table = sel.Tables.Add(sel.Range, rows, columns);
                table.Select();
                applyStyle?.Invoke(table);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{errorMessage}：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region 设置表格样式方法
        public void SetCurrentTableToThreeLineStyle()
        {
            SetCurrentTableStyle(SetThreeLineTableStyle);
        }

        public void SetCurrentTableToGBStyle()
        {
            SetCurrentTableStyle(SetGBTableStyle);
        }

        public void SetCurrentTableToNoBorderStyle()
        {
            SetCurrentTableStyle(SetNoBorderTableStyle);
        }

        private void SetCurrentTableStyle(Action<Word.Table> applyStyle)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            if (sel == null || sel.Tables.Count == 0)
            {
                MessageBox.Show("请将光标放在表格内！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Word.Table table = sel.Tables[1];
                applyStyle(table);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置表格样式失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        public void SetNoBorderTableStyle(Word.Table table)
        {
            if (table == null)
                return;

            try
            {
                // 清除所有边框
                ClearAllBorders(table);

                // 设置全部居中（水平和垂直）
                table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                // 关闭表格尺寸重调功能（仅无框线表格需要关闭，其他样式保持默认或启用）
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed);

                // 设置单元格左右边距为0
                SetCellPaddingToZero(table);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置无框线表格失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void SetGBTableStyle(Word.Table table)
        {
            if (table == null)
                return;

            // 获取表格的行列范围
            var range = GetTableRange(table);
            int firstRowIndex = range.Item1;
            int lastRowIndex = range.Item2;
            int firstColIndex = range.Item3;
            int lastColIndex = range.Item4;

            // 先设置所有内部框线为0.75磅
            foreach (Word.Cell cell in table.Range.Cells)
            {
                int rowIdx = cell.RowIndex;
                int colIdx = cell.ColumnIndex;

                // 内部水平线：0.75磅（排除最后一行）
                if (rowIdx < lastRowIndex)
                {
                    cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    cell.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth075pt;
                }

                // 内部垂直线：0.75磅（排除第一列）
                if (colIdx > firstColIndex)
                {
                    cell.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    cell.Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth075pt;
                }
            }

            // 设置外边框为1.5磅（覆盖内部框线）
            foreach (Word.Cell cell in table.Range.Cells)
            {
                int rowIdx = cell.RowIndex;
                int colIdx = cell.ColumnIndex;

                // 顶部外边框：1.5磅
                if (rowIdx == firstRowIndex)
                {
                    cell.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    cell.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                }

                // 底部外边框：1.5磅
                if (rowIdx == lastRowIndex)
                {
                    cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    cell.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                }

                // 左侧外边框：1.5磅
                if (colIdx == firstColIndex)
                {
                    cell.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    cell.Borders[Word.WdBorderType.wdBorderLeft].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                }

                // 右侧外边框：1.5磅
                if (colIdx == lastColIndex)
                {
                    cell.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    cell.Borders[Word.WdBorderType.wdBorderRight].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                }
            }

            // 获取第一行用于设置标题栏
            Word.Row firstRow = table.Rows[firstRowIndex];

            // 设置标题栏（第一行）的下边框为1.5磅（如果第一行不是最后一行）
            if (firstRowIndex != lastRowIndex)
            {
                foreach (Word.Cell cell in firstRow.Cells)
                {
                    cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    cell.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth150pt;
                }
            }

            // 设置标题栏内容居中对齐
            foreach (Word.Cell cell in firstRow.Cells)
            {
                cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
        }

        public void SetThreeLineTableStyle(Word.Table table)
        {
            if (table == null)
                return;

            // 获取表格的行范围
            var rowRange = GetTableRowRange(table);
            int firstRowIndex = rowRange.Item1;
            int lastRowIndex = rowRange.Item2;

            // 清除所有单元格的所有边框
            ClearAllBorders(table);

            // 设置三线表边框
            Word.Row firstRow = table.Rows[firstRowIndex];
            // 第一条线：表格顶部（第一行顶部）1.5磅
            SetRowBorder(firstRow, Word.WdBorderType.wdBorderTop, Word.WdLineWidth.wdLineWidth150pt);

            // 第二条线：标题栏下方（第一行底部）0.75磅（仅当表格有多行时）
            if (firstRowIndex != lastRowIndex)
            {
                SetRowBorder(firstRow, Word.WdBorderType.wdBorderBottom, Word.WdLineWidth.wdLineWidth075pt);
            }

            // 第三条线：表格底部（最后一行底部）1.5磅
            Word.Row lastRow = table.Rows[lastRowIndex];
            SetRowBorder(lastRow, Word.WdBorderType.wdBorderBottom, Word.WdLineWidth.wdLineWidth150pt);

            // 设置表格对齐格式（字体、行距、宽度、自动调整使用默认值）
            SetTableAlignment(table);
        }

        #region 辅助方法 - 表格样式设置
        private void ClearAllBorders(Word.Table table)
        {
            Word.WdBorderType[] borderTypes = new[]
            {
                Word.WdBorderType.wdBorderLeft,
                Word.WdBorderType.wdBorderRight,
                Word.WdBorderType.wdBorderTop,
                Word.WdBorderType.wdBorderBottom
            };

            foreach (Word.Cell cell in table.Range.Cells)
            {
                foreach (Word.WdBorderType borderType in borderTypes)
                {
                    try
                    {
                        cell.Borders[borderType].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                    }
                    catch (Exception ex)
                    {
                        // 如果某个边框设置失败，继续处理其他边框
                        System.Diagnostics.Debug.WriteLine($"设置边框失败: {ex.Message}");
                    }
                }
            }
        }

        private void SetTableAlignment(Word.Table table)
        {
            table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
            table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
        }

        private void SetCellPaddingToZero(Word.Table table)
        {
            try
            {
                // 使用表格级别的默认单元格边距设置
                table.LeftPadding = 0f;
                table.RightPadding = 0f;
            }
            catch
            {
                // 如果表格级别的边距设置失败，尝试逐个单元格设置
                try
                {
                    foreach (Word.Cell cell in table.Range.Cells)
                    {
                        cell.LeftPadding = 0f;
                        cell.RightPadding = 0f;
                    }
                }
                catch (Exception ex)
                {
                    // 如果单元格级别的边距设置也失败，记录错误但继续执行
                    System.Diagnostics.Debug.WriteLine($"设置单元格边距失败: {ex.Message}");
                }
            }
        }

        private Tuple<int, int> GetTableRowRange(Word.Table table)
        {
            int firstRowIndex = int.MaxValue;
            int lastRowIndex = int.MinValue;

            foreach (Word.Cell cell in table.Range.Cells)
            {
                if (cell.RowIndex < firstRowIndex)
                    firstRowIndex = cell.RowIndex;
                if (cell.RowIndex > lastRowIndex)
                    lastRowIndex = cell.RowIndex;
            }

            return Tuple.Create(firstRowIndex, lastRowIndex);
        }

        private Tuple<int, int, int, int> GetTableRange(Word.Table table)
        {
            int firstRowIndex = int.MaxValue;
            int lastRowIndex = int.MinValue;
            int firstColIndex = int.MaxValue;
            int lastColIndex = int.MinValue;

            foreach (Word.Cell cell in table.Range.Cells)
            {
                if (cell.RowIndex < firstRowIndex)
                    firstRowIndex = cell.RowIndex;
                if (cell.RowIndex > lastRowIndex)
                    lastRowIndex = cell.RowIndex;
                if (cell.ColumnIndex < firstColIndex)
                    firstColIndex = cell.ColumnIndex;
                if (cell.ColumnIndex > lastColIndex)
                    lastColIndex = cell.ColumnIndex;
            }

            return Tuple.Create(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex);
        }

        private void SetRowBorder(Word.Row row, Word.WdBorderType borderType, Word.WdLineWidth lineWidth)
        {
            foreach (Word.Cell cell in row.Cells)
            {
                cell.Borders[borderType].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                cell.Borders[borderType].LineWidth = lineWidth;
            }
        }
        #endregion

        public void InsertNRows()
        {
            var sel = Globals.ThisAddIn.Application.Selection;
            if (!ValidateTableSelection(sel))
                return;

            InsertOperationResult result = ShowInsertDialog("请输入要插入的行数：", "插入行", "往上插入", "往下插入");
            if (result.Cancelled)
                return;

            Word.Table table = sel.Tables[1];
            Word.Row refRow = GetReferenceRow(sel, table);

            try
            {
                for (int i = 0; i < result.Count; i++)
                {
                    if (result.Direction == InsertDirection.Before)
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

        public void InsertNColumns()
        {
            var sel = Globals.ThisAddIn.Application.Selection;
            if (!ValidateTableSelection(sel))
                return;

            InsertOperationResult result = ShowInsertDialog("请输入要插入的列数：", "插入列", "往左插入", "往右插入");
            if (result.Cancelled)
                return;

            Word.Table table = sel.Tables[1];
            Word.Column refCol = GetReferenceColumn(sel, table);

            try
            {
                for (int i = 0; i < result.Count; i++)
                {
                    if (result.Direction == InsertDirection.Before)
                        table.Columns.Add(refCol);
                    else
                        table.Columns.Add(refCol.Next);
                }
                
                // 自动调整列宽，确保表格不超出页面
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
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

        private class InsertOperationResult
        {
            public int Count { get; set; }
            public InsertDirection Direction { get; set; }
            public bool Cancelled { get; set; }
        }

        private enum InsertDirection
        {
            Before,  // 之前（上/左）
            After    // 之后（下/右）
        }

        private InsertOperationResult ShowInsertDialog(string countPrompt, string title, string beforeButtonText, string afterButtonText)
        {
            Form insertForm = new Form();
            Label promptLabel = new Label();
            TextBox inputTextBox = new TextBox();
            Button beforeButton = new Button();
            Button afterButton = new Button();
            Button cancelButton = new Button();

            insertForm.Text = title;
            insertForm.FormBorderStyle = FormBorderStyle.FixedDialog;
            insertForm.MaximizeBox = false;
            insertForm.MinimizeBox = false;
            insertForm.StartPosition = FormStartPosition.CenterScreen;
            insertForm.ShowInTaskbar = false;

            // 设置窗口宽度（足够宽以显示完整文本）
            int formWidth = 450;
            int padding = 30;
            int inputWidth = 350;

            // 设置提示标签
            promptLabel.Text = countPrompt;
            promptLabel.Font = new Font("微软雅黑", 12F, FontStyle.Bold);
            promptLabel.Location = new Point(padding, 30);
            promptLabel.AutoSize = true;
            promptLabel.MaximumSize = new Size(formWidth - padding * 2, 0);

            // 设置输入框（在label和按钮之间居中）
            inputTextBox.Font = new Font("微软雅黑", 14F, FontStyle.Regular);
            inputTextBox.Text = "1";
            inputTextBox.Width = inputWidth;
            inputTextBox.Height = 35;

            // 设置按钮
            int buttonWidth = 95;
            int buttonHeight = 35;
            int buttonSpacing = 10;
            int totalButtonWidth = buttonWidth * 3 + buttonSpacing * 2;
            // 整体往左移动12像素，使输入框和按钮更居中
            int horizontalOffset = 12;
            int buttonStartX = (formWidth - totalButtonWidth) / 2 - horizontalOffset;
            int inputBoxX = (formWidth - inputWidth) / 2 - horizontalOffset;

            beforeButton.Text = beforeButtonText;
            beforeButton.Font = new Font("微软雅黑", 11F, FontStyle.Bold);
            beforeButton.Size = new Size(buttonWidth, buttonHeight);
            beforeButton.Tag = InsertDirection.Before;
            beforeButton.Click += (s, e) => {
                insertForm.DialogResult = DialogResult.OK;
                insertForm.Tag = InsertDirection.Before;
                insertForm.Close();
            };

            afterButton.Text = afterButtonText;
            afterButton.Font = new Font("微软雅黑", 11F, FontStyle.Bold);
            afterButton.Size = new Size(buttonWidth, buttonHeight);
            afterButton.Tag = InsertDirection.After;
            afterButton.Click += (s, e) => {
                insertForm.DialogResult = DialogResult.OK;
                insertForm.Tag = InsertDirection.After;
                insertForm.Close();
            };

            cancelButton.Text = "取消";
            cancelButton.Font = new Font("微软雅黑", 11F, FontStyle.Bold);
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Size = new Size(buttonWidth, buttonHeight);

            // 设置窗体大小（按钮距离底部30）
            insertForm.Width = formWidth;
            int formHeight = 30 + promptLabel.Height + 100 + inputTextBox.Height + 50 + buttonHeight + 30;
            insertForm.Height = formHeight;

            // 计算输入框位置（在label和按钮之间居中）
            int availableHeight = formHeight - 30 - promptLabel.Height - 50 - buttonHeight;
            int inputBoxY = 30 + promptLabel.Height + (availableHeight - inputTextBox.Height) / 2;
            inputTextBox.Location = new Point(inputBoxX, inputBoxY);

            // 计算按钮位置（距离底部30）
            int buttonY = formHeight - 70 - buttonHeight;
            beforeButton.Location = new Point(buttonStartX, buttonY);
            afterButton.Location = new Point(beforeButton.Right + buttonSpacing, buttonY);
            cancelButton.Location = new Point(afterButton.Right + buttonSpacing, buttonY);

            // 添加控件
            insertForm.Controls.Add(promptLabel);
            insertForm.Controls.Add(inputTextBox);
            insertForm.Controls.Add(beforeButton);
            insertForm.Controls.Add(afterButton);
            insertForm.Controls.Add(cancelButton);

            insertForm.CancelButton = cancelButton;
            insertForm.AcceptButton = afterButton; // 默认选中"往右/下"按钮，方便直接按回车

            // 输入框回车键事件，触发默认按钮
            inputTextBox.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    afterButton.PerformClick();
                }
            };

            // 默认焦点在"往右/下"按钮上，方便直接按回车
            insertForm.Load += (s, e) => {
                inputTextBox.SelectAll();
                afterButton.Focus();
            };

            InsertOperationResult result = new InsertOperationResult();

            if (insertForm.ShowDialog() == DialogResult.OK)
            {
                // 验证输入的数量
                if (string.IsNullOrWhiteSpace(inputTextBox.Text) || !int.TryParse(inputTextBox.Text, out int count) || count <= 0)
            {
                MessageBox.Show("请输入有效的正整数！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    result.Cancelled = true;
                    return result;
        }

                result.Count = count;
                result.Direction = (InsertDirection)insertForm.Tag;
                result.Cancelled = false;
            }
            else
            {
                result.Cancelled = true;
            }

            return result;
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