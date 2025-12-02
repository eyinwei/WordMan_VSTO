using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
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

        /// <summary>
        /// 插入N列
        /// </summary>
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

        /// <summary>
        /// 插入操作结果
        /// </summary>
        private class InsertOperationResult
        {
            public int Count { get; set; }
            public InsertDirection Direction { get; set; }
            public bool Cancelled { get; set; }
        }

        /// <summary>
        /// 插入方向
        /// </summary>
        private enum InsertDirection
        {
            Before,  // 之前（上/左）
            After    // 之后（下/右）
        }

        /// <summary>
        /// 显示插入操作的合并对话框（包含数量输入和方向选择）
        /// </summary>
        /// <param name="countPrompt">数量提示文本</param>
        /// <param name="title">窗口标题</param>
        /// <param name="beforeButtonText">第一个按钮文本（上/左）</param>
        /// <param name="afterButtonText">第二个按钮文本（下/右，默认选中）</param>
        /// <returns>插入操作结果</returns>
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
            promptLabel.Font = new Font("微软雅黑", 12F, FontStyle.Regular);
            promptLabel.Location = new Point(padding, padding);
            promptLabel.AutoSize = true;
            promptLabel.MaximumSize = new Size(formWidth - padding * 2, 0);

            // 设置输入框
            inputTextBox.Font = new Font("微软雅黑", 14F, FontStyle.Regular);
            inputTextBox.Text = "1";
            inputTextBox.Location = new Point(padding, promptLabel.Bottom + 25);
            inputTextBox.Width = inputWidth;
            inputTextBox.Height = 35;

            // 设置按钮
            int buttonWidth = 95;
            int buttonHeight = 35;
            int buttonY = inputTextBox.Bottom + 30;
            int buttonSpacing = 10;
            int totalButtonWidth = buttonWidth * 3 + buttonSpacing * 2;
            int buttonStartX = (formWidth - totalButtonWidth) / 2;

            beforeButton.Text = beforeButtonText;
            beforeButton.Font = new Font("微软雅黑", 11F, FontStyle.Regular);
            beforeButton.Size = new Size(buttonWidth, buttonHeight);
            beforeButton.Location = new Point(buttonStartX, buttonY);
            beforeButton.Tag = InsertDirection.Before;
            beforeButton.Click += (s, e) => {
                insertForm.DialogResult = DialogResult.OK;
                insertForm.Tag = InsertDirection.Before;
                insertForm.Close();
            };

            afterButton.Text = afterButtonText;
            afterButton.Font = new Font("微软雅黑", 11F, FontStyle.Regular);
            afterButton.Size = new Size(buttonWidth, buttonHeight);
            afterButton.Location = new Point(beforeButton.Right + buttonSpacing, buttonY);
            afterButton.Tag = InsertDirection.After;
            afterButton.Click += (s, e) => {
                insertForm.DialogResult = DialogResult.OK;
                insertForm.Tag = InsertDirection.After;
                insertForm.Close();
            };

            cancelButton.Text = "取消";
            cancelButton.Font = new Font("微软雅黑", 11F, FontStyle.Regular);
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Size = new Size(buttonWidth, buttonHeight);
            cancelButton.Location = new Point(afterButton.Right + buttonSpacing, buttonY);

            // 设置窗体大小
            insertForm.Width = formWidth;
            insertForm.Height = buttonY + buttonHeight + padding + 40;

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