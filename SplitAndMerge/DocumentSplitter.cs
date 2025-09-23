using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    /// <summary>
    /// 文档拆分器 - 提供Word文档按页拆分功能
    /// </summary>
    public class DocumentSplitter
    {
        private readonly Word.Application _wordApplication;

        /// <summary>
        /// 初始化文档拆分器
        /// </summary>
        /// <param name="wordApplication">Word应用程序实例</param>
        public DocumentSplitter(Word.Application wordApplication)
        {
            _wordApplication = wordApplication;
        }

        /// <summary>
        /// 显示文档拆分对话框并执行拆分操作
        /// </summary>
        public void ShowSplitDialog()
        {

            var openFileDialog = new OpenFileDialog
            {
                Title = "选择要拆分的Word文档",
                Filter = "Word文档 (*.docx)|*.docx|Word文档 (*.doc)|*.doc|所有文件 (*.*)|*.*",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filePath = openFileDialog.FileName;
                var result = MessageBox.Show("确认进行逐页拆分？", "文档拆分", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    ExecuteSplit(filePath);
                }
            }
        }


        /// <summary>
        /// 执行拆分操作
        /// </summary>
        private void ExecuteSplit(string filePath)
        {
            Word.Document document = null;
            
            try
            {
                document = _wordApplication.Documents.Open(filePath);
                ValidateDocumentState(document);
                
                SplitDocumentByPages(document);
            }
            catch (Exception ex)
            {
                ShowMessage($"文档拆分失败：{ex.Message}", "错误", MessageBoxIcon.Error);
            }
            finally
            {
                SafeCloseDocument(document);
            }
        }

        /// <summary>
        /// 显示消息框
        /// </summary>
        private void ShowMessage(string message, string title, MessageBoxIcon icon = MessageBoxIcon.Information)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, icon);
        }


        /// <summary>
        /// 按页拆分文档
        /// </summary>
        private void SplitDocumentByPages(Word.Document document)
        {
            string outputFolder = PrepareSplitEnvironment(document);
            int totalPages = document.Range().Information[WdInformation.wdNumberOfPagesInDocument];
            string baseFileName = Path.GetFileNameWithoutExtension(document.FullName);

            for (int pageNumber = 1; pageNumber <= totalPages; pageNumber++)
            {
                SplitSinglePageWithRetry(document, pageNumber, outputFolder, baseFileName);
                if (pageNumber % 3 == 0) { GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); }
            }

            ShowSplitResult(false, totalPages, outputFolder);
        }



        /// <summary>
        /// 显示拆分结果
        /// </summary>
        private void ShowSplitResult(bool isCancelled, int fileCount, string outputFolder)
        {
            if (isCancelled)
            {
                MessageBox.Show("文档拆分已取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show($"文档已成功拆分为 {fileCount} 个文件，保存在：{outputFolder}", "拆分完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        /// <summary>
        /// 获取指定页的范围
        /// </summary>
        private Word.Range GetPageRange(Word.Document document, int startPage, int endPage)
        {
            Word.Range startRange = null;
            Word.Range endRange = null;
            Word.Range resultRange = null;
            
            try
            {
                int totalPages = document.Range().Information[WdInformation.wdNumberOfPagesInDocument];
                if (startPage < 1 || endPage < 1 || startPage > totalPages || endPage > totalPages)
                    throw new Exception($"页码范围无效：第{startPage}-{endPage}页（总页数：{totalPages}）");
                
                if (startPage > endPage)
                    throw new Exception($"起始页不能大于结束页：第{startPage}-{endPage}页");
                
                // 尝试GoTo方法
                try
                {
                    startRange = document.GoTo(What: WdGoToItem.wdGoToPage, Which: WdGoToDirection.wdGoToAbsolute, Count: startPage);
                    if (startRange != null)
                    {
                        int startPos = startRange.Start;
                        int endPos = GetPageEndPosition(document, endPage, totalPages);
                        
                        resultRange = document.Range();
                        resultRange.SetRange(startPos, Math.Max(startPos, endPos));
                        
                        if (resultRange.Start < resultRange.End) return resultRange;
                    }
                }
                catch { }
                finally
                {
                    ReleaseComObject(startRange);
                    ReleaseComObject(endRange);
                }
                
                // 备用方法：字符位置估算
                return GetPageRangeByCharacterPosition(document, startPage, endPage, totalPages);
            }
            catch (Exception ex)
            {
                ReleaseComObject(resultRange);
                throw new Exception($"获取第{startPage}-{endPage}页范围失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 获取页面结束位置
        /// </summary>
        private int GetPageEndPosition(Word.Document document, int endPage, int totalPages)
        {
            if (endPage < totalPages)
            {
                var nextPageRange = document.GoTo(What: WdGoToItem.wdGoToPage, Which: WdGoToDirection.wdGoToAbsolute, Count: endPage + 1);
                if (nextPageRange != null)
                {
                    int endPos = nextPageRange.Start - 1;
                    ReleaseComObject(nextPageRange);
                    return endPos;
                }
            }
            return document.Range().End;
        }

        /// <summary>
        /// 使用字符位置估算获取页面范围
        /// </summary>
        private Word.Range GetPageRangeByCharacterPosition(Word.Document document, int startPage, int endPage, int totalPages)
        {
            int totalCharacters = document.Range().End;
            int charactersPerPage = Math.Max(1, totalCharacters / totalPages);
            
            int startPos = Math.Max(0, (startPage - 1) * charactersPerPage);
            int endPos = Math.Min(totalCharacters, endPage * charactersPerPage);
            
            if (startPos >= endPos) startPos = Math.Max(0, endPos - 50);
            
            var range = document.Range();
            range.SetRange(startPos, endPos);
            return range;
        }

        /// <summary>
        /// 验证文档状态
        /// </summary>
        private void ValidateDocumentState(Word.Document doc)
        {
            if (doc == null) throw new Exception("文档对象为空");
            
            if (doc.ProtectionType != WdProtectionType.wdNoProtection)
                throw new Exception("文档被保护，无法进行拆分操作");
            
            int totalPages = doc.Range().Information[WdInformation.wdNumberOfPagesInDocument];
            if (totalPages <= 0) throw new Exception("文档没有有效页面");
        }

        /// <summary>
        /// 带重试机制的拆分单页方法
        /// </summary>
        private void SplitSinglePageWithRetry(Word.Document document, int pageNumber, string outputFolder, string baseFileName)
        {
            const int maxRetries = 3;
            Exception lastException = null;
            
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    if (attempt > 1) System.Threading.Thread.Sleep(500);
                    
                    SplitPageRange(document, pageNumber, pageNumber, outputFolder, baseFileName);
                    return;
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    if (attempt < maxRetries)
                    {
                        try
                        {
                            _wordApplication.ScreenUpdating = false;
                            _wordApplication.ScreenUpdating = true;
                        }
                        catch { }
                    }
                }
            }
            
            throw new Exception($"拆分第{pageNumber}页失败（已重试{maxRetries}次）：{lastException?.Message}");
        }


        /// <summary>
        /// 复制页面设置
        /// </summary>
        private void CopyPageSetup(Word.Document sourceDoc, Word.Document targetDoc, int pageNum = 1)
        {
            try
            {
                var sourceSection = GetPageSection(sourceDoc, pageNum) ?? sourceDoc.Sections[1];
                var targetSection = targetDoc.Sections[1];
                var sourcePageSetup = sourceSection.PageSetup;
                var targetPageSetup = targetSection.PageSetup;
                
                targetPageSetup.TopMargin = sourcePageSetup.TopMargin;
                targetPageSetup.BottomMargin = sourcePageSetup.BottomMargin;
                targetPageSetup.LeftMargin = sourcePageSetup.LeftMargin;
                targetPageSetup.RightMargin = sourcePageSetup.RightMargin;
                targetPageSetup.HeaderDistance = sourcePageSetup.HeaderDistance;
                targetPageSetup.FooterDistance = sourcePageSetup.FooterDistance;
                targetPageSetup.PageWidth = sourcePageSetup.PageWidth;
                targetPageSetup.PageHeight = sourcePageSetup.PageHeight;
                targetPageSetup.Orientation = sourcePageSetup.Orientation;
                targetPageSetup.PaperSize = sourcePageSetup.PaperSize;
                targetPageSetup.Gutter = sourcePageSetup.Gutter;
                targetPageSetup.MirrorMargins = sourcePageSetup.MirrorMargins;
                targetPageSetup.TwoPagesOnOne = sourcePageSetup.TwoPagesOnOne;
                targetPageSetup.BookFoldPrinting = sourcePageSetup.BookFoldPrinting;
                targetPageSetup.BookFoldRevPrinting = sourcePageSetup.BookFoldRevPrinting;
                
                CopyHeadersAndFooters(sourceSection, targetSection);
            }
            catch { }
        }

        /// <summary>
        /// 获取指定页面所属的节
        /// </summary>
        private Word.Section GetPageSection(Word.Document doc, int pageNum)
        {
            try
            {
                var pageRange = doc.GoTo(What: WdGoToItem.wdGoToPage, Which: WdGoToDirection.wdGoToAbsolute, Count: pageNum);
                if (pageRange == null) return null;
                
                var section = pageRange.Sections[1];
                ReleaseComObject(pageRange);
                return section;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 复制页眉页脚
        /// </summary>
        private void CopyHeadersAndFooters(Word.Section sourceSection, Word.Section targetSection)
        {
            try
            {
                var sourceHeader = sourceSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                var targetHeader = targetSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                if (sourceHeader.Range.Text.Trim().Length > 0)
                {
                    sourceHeader.Range.Copy();
                    targetHeader.Range.Paste();
                }
                
                var sourceFooter = sourceSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                var targetFooter = targetSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                if (sourceFooter.Range.Text.Trim().Length > 0)
                {
                    sourceFooter.Range.Copy();
                    targetFooter.Range.Paste();
                }
            }
            catch { }
        }

        /// <summary>
        /// 清理文档末尾的空白内容
        /// </summary>
        private void CleanTrailingContent(Word.Document doc)
        {
            try
            {
                var docRange = doc.Range();
                if (docRange.End <= 1) return;
                
                var endRange = doc.Range();
                endRange.SetRange(Math.Max(0, docRange.End - 100), docRange.End);
                
                string trailingText = endRange.Text;
                bool hasOnlyWhitespace = trailingText.Trim().Length == 0 || 
                                       trailingText.All(c => char.IsWhiteSpace(c) || c == '\r' || c == '\n' || c == '\f');
                
                if (hasOnlyWhitespace)
                {
                    int lastNonWhitespace = -1;
                    for (int i = docRange.End - 1; i >= 0; i--)
                    {
                        var charRange = doc.Range(i, i + 1);
                        if (!string.IsNullOrWhiteSpace(charRange.Text) && charRange.Text != "\r" && charRange.Text != "\n")
                        {
                            lastNonWhitespace = i + 1;
                            break;
                        }
                        ReleaseComObject(charRange);
                    }
                    
                    if (lastNonWhitespace > 0 && lastNonWhitespace < docRange.End)
                    {
                        var deleteRange = doc.Range(lastNonWhitespace, docRange.End);
                        deleteRange.Delete();
                        ReleaseComObject(deleteRange);
                    }
                }
                
                ReleaseComObject(endRange);
            }
            catch { }
        }


        /// <summary>
        /// 安全释放COM对象
        /// </summary>
        private void ReleaseComObject(object comObject)
        {
            try
            {
                if (comObject != null)
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch { }
            finally
            {
                comObject = null;
            }
        }

        /// <summary>
        /// 安全关闭并释放文档对象
        /// </summary>
        private void SafeCloseDocument(Word.Document doc)
        {
            try
            {
                if (doc != null)
                {
                    doc.Close();
                    ReleaseComObject(doc);
                }
            }
            catch { }
        }

        /// <summary>
        /// 准备拆分环境
        /// </summary>
        private string PrepareSplitEnvironment(Word.Document sourceDoc)
        {
            string basePath = Path.GetDirectoryName(sourceDoc.FullName);
            string fileName = Path.GetFileNameWithoutExtension(sourceDoc.FullName);
            string splitFolder = Path.Combine(basePath, fileName + "_拆分");
            
            if (string.IsNullOrEmpty(basePath) || !Directory.Exists(basePath))
                throw new Exception("源文档路径无效，无法创建拆分文件夹");
            
            if (!IsPathWritable(basePath))
                throw new Exception("目标路径不可写，请检查权限设置");
            
            if (!Directory.Exists(splitFolder))
            {
                Directory.CreateDirectory(splitFolder);
            }
            
            if (!Directory.Exists(splitFolder))
                throw new Exception("无法创建拆分文件夹，请检查权限设置");
            
            return splitFolder;
        }

        /// <summary>
        /// 检查路径是否可写
        /// </summary>
        private bool IsPathWritable(string path)
        {
            try
            {
                string testFile = Path.Combine(path, "test_write.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 设置页面方向
        /// </summary>
        private void SetPageOrientation(Word.Document sourceDocument, Word.Document targetDocument, int pageNumber)
        {
            try
            {
                var sourceSection = GetPageSection(sourceDocument, pageNumber) ?? sourceDocument.Sections[1];
                var sourcePageSetup = sourceSection.PageSetup;
                var targetPageSetup = targetDocument.Sections[1].PageSetup;
                
                targetPageSetup.Orientation = sourcePageSetup.Orientation;
                targetPageSetup.PageWidth = sourcePageSetup.PageWidth;
                targetPageSetup.PageHeight = sourcePageSetup.PageHeight;
                targetPageSetup.PaperSize = sourcePageSetup.PaperSize;
                
                targetDocument.Range().Font.Reset();
            }
            catch { }
        }

        /// <summary>
        /// 拆分页面到新文档
        /// </summary>
        private void SplitPageRange(Word.Document sourceDocument, int startPage, int endPage, string outputFolder, string baseFileName)
        {
            Word.Range pageRange = null;
            Word.Document newDocument = null;
            Word.Selection selection = null;
            
            try
            {
                if (sourceDocument == null) throw new Exception("源文档为空");
                if (startPage < 1 || endPage < 1 || startPage > endPage)
                    throw new Exception($"页码范围无效：第{startPage}-{endPage}页");
                
                string rangeDescription = startPage == endPage ? $"第{startPage}页" : $"第{startPage}-{endPage}页";
                
                pageRange = GetPageRange(sourceDocument, startPage, endPage);
                if (pageRange == null || pageRange.Start >= pageRange.End)
                    throw new Exception($"{rangeDescription}内容为空或无效");

                pageRange.Copy();
                
                newDocument = _wordApplication.Documents.Add();
                if (newDocument == null) throw new Exception("无法创建新文档");
                
                SetPageOrientation(sourceDocument, newDocument, startPage);
                
                selection = newDocument.Application.Selection;
                if (selection == null) throw new Exception("无法获取选择对象");
                selection.Paste();
                
                try
                {
                    CopyPageSetup(sourceDocument, newDocument, startPage);
                }
                catch { }
                
                CleanTrailingContent(newDocument);
                
                string fileName = $"{baseFileName}_{rangeDescription}.docx";
                string outputFilePath = Path.Combine(outputFolder, fileName);
                
                newDocument.SaveAs2(FileName: outputFilePath, FileFormat: WdSaveFormat.wdFormatXMLDocument);
            }
            catch (Exception ex)
            {
                throw new Exception($"拆分第{startPage}-{endPage}页失败：{ex.Message}");
            }
            finally
            {
                ReleaseComObject(selection);
                ReleaseComObject(pageRange);
                SafeCloseDocument(newDocument);
            }
        }

    }
}

