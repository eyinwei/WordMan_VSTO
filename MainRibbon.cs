﻿using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.VisualBasic;
using Word = Microsoft.Office.Interop.Word;
using WordMan.SplitAndMerge;
using WordMan.MultiLevel;
using static WordMan.CaptionManager;
using WordMan;

namespace WordMan
{
    public partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private ImageProcessor imageProcessor = new ImageProcessor();
        private TextProcessor textProcessor = new TextProcessor();
        private TableProcessor tableProcessor = new TableProcessor();

        private void 去除断行_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.RemoveLineBreaks();
        }
        private void 去除空格_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.RemoveSpaces();
        }
        private void 去除空行_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.RemoveEmptyLines();
        }

        private void 英标转中标_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.ConvertEnglishToChinesePunctuation();
        }
        private void 中标转英标_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.ConvertChineseToEnglishPunctuation();
        }

        private void 自动加空格_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.AutoAddSpaces();
        }

        private void 缩进2字符_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.IndentTwoCharacters();
        }

        private void 去除缩进_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.RemoveIndent();
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

        private void 创建三线表_Click(object sender, RibbonControlEventArgs e)
        {
            tableProcessor.CreateThreeLineTable();
        }

        private void 设为三线_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            if (sel == null || sel.Tables.Count == 0)
                return;

            Word.Table table = sel.Tables[1];
            tableProcessor.SetThreeLineTable(table);
        }

        private void 插入N行_Click(object sender, RibbonControlEventArgs e)
        {
            tableProcessor.InsertNRows();
        }

        private void 插入N列_Click(object sender, RibbonControlEventArgs e)
        {
            tableProcessor.InsertNColumns();
        }

        // 公式编号样式相关枚举和变量
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
            // 保存原始选择位置
            int originalStart = 0;
            int originalEnd = 0;
            
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                
                // 保存当前光标位置
                originalStart = sel.Start;
                originalEnd = sel.End;
                
                Word.Paragraph para = sel.Paragraphs[1];

                // 1. 选择段落内容（不包含段落标记）并剪切
                Word.Range contentRange = para.Range.Duplicate;
                contentRange.End = contentRange.End - 1; // 排除段落标记
                contentRange.Cut();

                // 2. 删除当前段落
                para.Range.Delete();

                // 3. 创建并配置表格
                Word.Table table = CaptionManager.CreateFormulaTable(sel, app);

                // 4. 粘贴公式内容到第二列
                table.Cell(1, 2).Range.Paste();

                // 5. 插入公式编号到第三列
                CaptionManager.InsertFormulaNumber(table, sel, CurrentStyle);

                // 6. 将光标移动到表格第二列
                table.Cell(1, 2).Range.Select();
            }
            catch (Exception ex)
            {
                // 如果出错，恢复原始选择
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    var sel = app.Selection;
                    sel.SetRange(originalStart, originalEnd);
                }
                catch { }
                
                MessageBox.Show($"公式编号插入失败：{ex.Message}\n\n请确保光标位于包含公式的段落中。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 图片编号样式相关枚举和变量
        private PictureNumberStyle CurrentPicStyle = PictureNumberStyle.Arabic;

        // 表格编号样式相关枚举和变量
        private TableNumberStyle CurrentTableStyle = TableNumberStyle.Arabic;

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
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.Selection;
                var doc = app.ActiveDocument;

                HashSet<int> handledParagraphs = new HashSet<int>();
                List<Word.Paragraph> targetParas = new List<Word.Paragraph>();

                // 选区有图片
                try
                {
                    foreach (Word.InlineShape ils in sel.Range.InlineShapes)
                    {
                        var para = ils.Range.Paragraphs.First;
                        if (!handledParagraphs.Contains(para.Range.Start))
                        {
                            targetParas.Add(para);
                            handledParagraphs.Add(para.Range.Start);
                        }
                    }
                }
                catch { }
                
                try
                {
                    foreach (Word.Shape s in sel.Range.ShapeRange)
                    {
                        var para = s.Anchor.Paragraphs.First;
                        if (!handledParagraphs.Contains(para.Range.Start))
                        {
                            targetParas.Add(para);
                            handledParagraphs.Add(para.Range.Start);
                        }
                    }
                }
                catch { }

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
                    CaptionManager.InsertPictureCaption(targetParas[i], CurrentPicStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"图片编号插入失败：{ex.Message}\n\n请确保光标位于包含图片的段落中。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            try
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
                    CaptionManager.InsertTableCaption(targetTables[i], CurrentTableStyle);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"表格编号插入失败：{ex.Message}\n\n请确保光标位于包含表格的段落中。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void 宽度刷_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.WidthBrush_Click(sender, e, 宽度刷);
        }

        private void 高度刷_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.HeightBrush_Click(sender, e, 高度刷);
        }

        // 在Cleanup方法中添加高度刷的清理
        public void Cleanup()
        {
            imageProcessor.Cleanup();
        }

        // 位图化按钮点击事件
        private void 位图化_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.ConvertToBitmap_Click(sender, e);
        }

        // 添加字段来存储原始位置和状态
        private Word.Range originalRange;
        private bool isCrossReferenceMode = false;
        private Microsoft.Office.Tools.Ribbon.RibbonToggleButton crossReferenceToggleButton;
        private System.Windows.Forms.Timer escKeyListener;

        private void 交叉引用_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取toggleButton
            crossReferenceToggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;

            // 如果已经在交叉引用模式，则退出
            if (isCrossReferenceMode)
            {
                ExitCrossReferenceMode();
                return;
            }

            try
            {
                // 记录当前光标位置
                originalRange = Globals.ThisAddIn.Application.Selection.Range;

                // 启用交叉引用模式
                isCrossReferenceMode = true;

                // 设置按钮为按下状态
                crossReferenceToggleButton.Checked = true;

                // 注册选择变化事件监听器
                Globals.ThisAddIn.Application.WindowSelectionChange += Application_WindowSelectionChange;

                // 初始化ESC键监听器
                InitializeEscKeyListener();

                // 更新状态栏提示
                Globals.ThisAddIn.Application.StatusBar = "交叉引用模式：请将光标移动到题注所在行，按ESC或再次点击按钮退出";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动交叉引用模式失败: {ex.Message}", "错误",
                               MessageBoxButtons.OK, MessageBoxIcon.Error);
                isCrossReferenceMode = false;
                crossReferenceToggleButton.Checked = false;
            }
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            // 只在交叉引用模式下处理
            if (!isCrossReferenceMode) return;

            // 获取当前选择位置
            Word.Range currentRange = Sel.Range;

            // 查找题注标签和编号
            CaptionManager.CaptionInfo captionInfo = CaptionManager.FindCaptionInfo(currentRange);

            if (captionInfo != null)
            {
                // 找到题注，插入交叉引用
                CaptionManager.InsertCrossReferenceAtOriginalPosition(originalRange, captionInfo);

                // 退出交叉引用模式
                ExitCrossReferenceMode();
            }
        }

        private void InitializeEscKeyListener()
        {
            // 初始化ESC键监听器
            escKeyListener = new System.Windows.Forms.Timer();
            escKeyListener.Interval = 100; // 100毫秒检查一次
            escKeyListener.Tick += EscKeyListener_Tick;
            escKeyListener.Start();
        }

        private void EscKeyListener_Tick(object sender, EventArgs e)
        {
            // 检查ESC键是否被按下
            if (isCrossReferenceMode && (GetAsyncKeyState(Keys.Escape) & 0x8000) != 0)
            {
                ExitCrossReferenceMode();
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(Keys vKey);

        private void ExitCrossReferenceMode()
        {
            // 退出交叉引用模式
            isCrossReferenceMode = false;

            // 取消注册事件监听器
            Globals.ThisAddIn.Application.WindowSelectionChange -= Application_WindowSelectionChange;

            // 停止ESC键监听器
            if (escKeyListener != null)
            {
                escKeyListener.Stop();
                escKeyListener.Dispose();
                escKeyListener = null;
            }

            // 清除状态栏提示
            Globals.ThisAddIn.Application.StatusBar = "";

            // 恢复按钮状态
            if (crossReferenceToggleButton != null)
            {
                crossReferenceToggleButton.Checked = false;
            }

            // 清除原始位置引用
            originalRange = null;
        }

        // 排版按钮点击事件
        private void TypesettingButton_Click(object sender, RibbonControlEventArgs e)
        {
            // 仅一行：调用任务窗格的静态方法，剩下的全由任务窗格自己处理
            TypesettingTaskPane.TriggerShowOrHide();
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

                // 调用Word的"另存为"对话框
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
        private void 版本_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/eyinwei/WordMan_VSTO");
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
        private void 上标_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;

                if (doc == null || doc.Fields == null)
                {
                    System.Windows.Forms.MessageBox.Show("未检测到文档或文档没有字段。");
                    return;
                }

                int refCount = 0;
                int otherCount = 0;
                string[] excludeKeywords = { "图", "表", "公式", "figure", "table", "equation",
                           "fig", "tab", "图表", "图片", "图形", "插图" };

                app.ScreenUpdating = false;

                // 使用Fields.Count避免枚举时修改集合的问题
                int fieldCount = doc.Fields.Count;

                for (int i = 1; i <= fieldCount; i++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Word.Field field = doc.Fields[i];

                        // 更准确的REF字段判断
                        if (field.Type == Microsoft.Office.Interop.Word.WdFieldType.wdFieldRef ||
                            field.Type == Microsoft.Office.Interop.Word.WdFieldType.wdFieldSequence)
                        {
                            string codeText = field.Code?.Text ?? "";
                            string resultText = field.Result?.Text ?? "";

                            // 检查是否需要排除（图表公式等）
                            string combinedText = (codeText + " " + resultText).ToLower();
                            bool isExcluded = excludeKeywords.Any(keyword =>
                                combinedText.Contains(keyword.ToLower()));

                            if (!isExcluded)
                            {
                                // 安全地设置上标
                                if (field.Result != null && field.Result.Font != null)
                                {
                                    field.Result.Font.Superscript = 1;
                                    refCount++;
                                }
                            }
                            else
                            {
                                otherCount++;
                            }
                        }
                        else
                        {
                            otherCount++;
                        }
                    }
                    catch (Exception fieldEx)
                    {
                        // 单个字段处理失败时继续处理其他字段
                        System.Diagnostics.Debug.WriteLine($"处理字段时出错: {fieldEx.Message}");
                        continue;
                    }
                }

                app.ScreenUpdating = true;

                // 显示处理结果
                System.Windows.Forms.MessageBox.Show(
                    $"处理完成：\n" +
                    $"• 参考文献引用: {refCount} 个（已设为上标）\n" +
                    $"• 其他字段: {otherCount} 个（未处理）",
                    "完成",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                try
                {
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                }
                catch { }

                System.Windows.Forms.MessageBox.Show($"处理过程中出现错误：{ex.Message}");
            }
        }

        private void 正常_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;

                if (doc == null || doc.Fields == null)
                {
                    System.Windows.Forms.MessageBox.Show("未检测到文档或文档没有字段。");
                    return;
                }

                int refCount = 0;
                int otherCount = 0;
                string[] excludeKeywords = { "图", "表", "公式", "figure", "table", "equation",
                           "fig", "tab", "图表", "图片", "图形", "插图" };

                app.ScreenUpdating = false;

                // 使用Fields.Count避免枚举时修改集合的问题
                int fieldCount = doc.Fields.Count;

                for (int i = 1; i <= fieldCount; i++)
                {
                    try
                    {
                        Microsoft.Office.Interop.Word.Field field = doc.Fields[i];

                        // 更准确的REF字段判断
                        if (field.Type == Microsoft.Office.Interop.Word.WdFieldType.wdFieldRef ||
                            field.Type == Microsoft.Office.Interop.Word.WdFieldType.wdFieldSequence)
                        {
                            string codeText = field.Code?.Text ?? "";
                            string resultText = field.Result?.Text ?? "";

                            // 检查是否需要排除（图表公式等）
                            string combinedText = (codeText + " " + resultText).ToLower();
                            bool isExcluded = excludeKeywords.Any(keyword =>
                                combinedText.Contains(keyword.ToLower()));

                            if (!isExcluded)
                            {
                                // 安全地设置正常格式
                                if (field.Result != null && field.Result.Font != null)
                                {
                                    field.Result.Font.Superscript = 0;
                                    refCount++;
                                }
                            }
                            else
                            {
                                otherCount++;
                            }
                        }
                        else
                        {
                            otherCount++;
                        }
                    }
                    catch (Exception fieldEx)
                    {
                        // 单个字段处理失败时继续处理其他字段
                        System.Diagnostics.Debug.WriteLine($"处理字段时出错: {fieldEx.Message}");
                        continue;
                    }
                }

                app.ScreenUpdating = true;

                // 显示处理结果
                System.Windows.Forms.MessageBox.Show(
                    $"处理完成：\n" +
                    $"• 参考文献引用: {refCount} 个（已设为正常格式）\n" +
                    $"• 其他字段: {otherCount} 个（未处理）",
                    "完成",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                try
                {
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                }
                catch { }

                System.Windows.Forms.MessageBox.Show($"处理过程中出现错误：{ex.Message}");
            }
        }

    }
}