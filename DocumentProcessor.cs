using System;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan
{
    public class DocumentProcessor
    {
        #region 域名处理
        public void HighlightFields(bool highlight)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("未检测到文档。");
                return;
            }

            foreach (Word.Field field in doc.Fields)
            {
                string code = field.Code.Text.Trim();
                Word.Range fieldResult = field.Result;
                string fieldText = fieldResult.Text;

                if (code.StartsWith("REF", StringComparison.OrdinalIgnoreCase) ||
                    code.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
                {
                    if (highlight)
                    {
                        if (fieldText.Contains("图"))
                        {
                            fieldResult.Font.Color = Word.WdColor.wdColorBlue;
                        }
                        else if (fieldText.Contains("表"))
                        {
                            fieldResult.Font.Color = Word.WdColor.wdColorGreen;
                        }
                        else if (fieldText.Contains("公式"))
                        {
                            fieldResult.Font.Color = Word.WdColor.wdColorRed;
                        }
                        else
                        {
                            fieldResult.Font.Color = Word.WdColor.wdColorBrown;
                        }
                    }
                    else
                    {
                        fieldResult.Font.Color = Word.WdColor.wdColorBlack;
                    }
                }
                else if (field.Type == Word.WdFieldType.wdFieldAddin &&
                         (code.Contains("EN.CITE") || code.Contains("EN.CITATION")))
                {
                    fieldResult.Font.Color = highlight ? Word.WdColor.wdColorGold : Word.WdColor.wdColorBlack;
                }
            }

            MessageBox.Show(highlight ? "交叉引用与文献引用已高亮！" : "交叉引用与文献引用已取消高亮！");
        }
        #endregion

        #region 文献编号处理
        public void SetFieldSuperscript(bool setSuperscript)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;

                if (doc == null || doc.Fields == null)
                {
                    MessageBox.Show("未检测到文档或文档没有字段。");
                    return;
                }

                int refCount = 0;
                int otherCount = 0;
                string[] excludeKeywords = { "图", "表", "公式", "figure", "table", "equation",
                           "fig", "tab", "图表", "图片", "图形", "插图" };

                app.ScreenUpdating = false;

                int fieldCount = doc.Fields.Count;
                for (int i = 1; i <= fieldCount; i++)
                {
                    try
                    {
                        Word.Field field = doc.Fields[i];

                        if (field.Type == Word.WdFieldType.wdFieldRef ||
                            field.Type == Word.WdFieldType.wdFieldSequence)
                        {
                            string codeText = field.Code?.Text ?? "";
                            string resultText = field.Result?.Text ?? "";
                            string combinedText = (codeText + " " + resultText).ToLower();
                            bool isExcluded = excludeKeywords.Any(keyword =>
                                combinedText.Contains(keyword.ToLower()));

                            if (!isExcluded)
                            {
                                if (field.Result != null && field.Result.Font != null)
                                {
                                    field.Result.Font.Superscript = setSuperscript ? 1 : 0;
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
                        System.Diagnostics.Debug.WriteLine($"处理字段时出错: {fieldEx.Message}");
                        continue;
                    }
                }

                app.ScreenUpdating = true;

                string statusText = setSuperscript ? "已设为上标" : "已设为正常格式";
                MessageBox.Show(
                    $"处理完成：\n" +
                    $"• 参考文献引用: {refCount} 个（{statusText}）\n" +
                    $"• 其他字段: {otherCount} 个（未处理）",
                    "完成",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                try
                {
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                }
                catch { }

                MessageBox.Show($"处理过程中出现错误：{ex.Message}");
            }
        }
        #endregion

        #region PDF处理
        public void ExportToPDF()
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;

            if (string.IsNullOrEmpty(doc.Path))
            {
                MessageBox.Show(
                    "请先保存文档，再导出为PDF。",
                    "提示",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                app.Dialogs[Word.WdWordDialog.wdDialogFileSaveAs].Show();
                return;
            }

            try
            {
                string docPath = doc.FullName;
                string directory = System.IO.Path.GetDirectoryName(docPath);
                string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(docPath);
                string pdfPath = System.IO.Path.Combine(directory, fileNameWithoutExt + ".pdf");

                doc.ExportAsFixedFormat(
                    pdfPath,
                    Word.WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false,
                    OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Range: Word.WdExportRange.wdExportAllDocument,
                    Item: Word.WdExportItem.wdExportDocumentContent,
                    CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                    DocStructureTags: true,
                    BitmapMissingFonts: true,
                    UseISO19005_1: false
                );

                var result = MessageBox.Show(
                    "成功导出为PDF！是否现在打开该PDF？",
                    "导出成功",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(pdfPath);
                    }
                    catch (Exception exOpen)
                    {
                        MessageBox.Show(
                            "打开PDF文件出错：" + exOpen.Message,
                            "错误",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "导出PDF失败：" + ex.Message,
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void ShowVersion()
        {
            System.Diagnostics.Process.Start("https://github.com/eyinwei/WordMan_VSTO");
        }
        #endregion

        #region 密级处理
        public void AddSecurityLevel(string levelText)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                var selection = app.Selection;

                RemoveSecurityLevelFromCurrentPage();

                int currentPage = selection.Information[Word.WdInformation.wdActiveEndPageNumber];

                var pageSetup = doc.PageSetup;
                float leftMargin = pageSetup.LeftMargin;
                float topMargin = pageSetup.TopMargin;

                selection.HomeKey(Word.WdUnits.wdLine, Word.WdMovementType.wdMove);

                var textBox = doc.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 100, 20);

                textBox.TextFrame.TextRange.Text = levelText;

                var textRange = textBox.TextFrame.TextRange;
                textRange.Font.Name = "黑体";
                textRange.Font.Size = 12;
                textRange.Font.Bold = 1;
                textRange.Font.Color = Word.WdColor.wdColorBlack;

                textBox.Line.Visible = MsoTriState.msoFalse;
                textBox.Fill.Visible = MsoTriState.msoFalse;
                textBox.Width = app.CentimetersToPoints(3.0f);
                textBox.Height = app.CentimetersToPoints(0.8f);

                textBox.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                textBox.Left = 0;

                textBox.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin;
                textBox.Top = -textBox.Height;

                textBox.WrapFormat.Type = Word.WdWrapType.wdWrapNone;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"添加密级标签失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void RemoveSecurityLevelFromCurrentPage()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;
                var selection = app.Selection;

                int currentPage = selection.Information[Word.WdInformation.wdActiveEndPageNumber];

                foreach (Word.Shape shape in doc.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoTextBox)
                    {
                        string text = shape.TextFrame.TextRange.Text.Trim();
                        if (text == "公开" || text == "内部★" || text.Contains("密级"))
                        {
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
                                shape.Delete();
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        private void RemoveSecurityLevel()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var doc = app.ActiveDocument;

                foreach (Word.Shape shape in doc.Shapes)
                {
                    if (shape.Type == MsoShapeType.msoTextBox)
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
            }
        }
        #endregion
    }
}

