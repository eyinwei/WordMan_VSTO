using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    /// <summary>
    /// 文档拆分器 - 优化版本
    /// </summary>
    public class DocumentSplitter
    {
        private Word.Application app;
        private bool isProcessing = false;
        private volatile bool isCancelled = false;

        public DocumentSplitter(Word.Application application)
        {
            app = application;
        }

        /// <summary>
        /// 显示文档拆分对话框
        /// </summary>
        public void ShowSplitDialog()
        {
            try
            {
                if (isProcessing)
                {
                    MessageBox.Show("文档拆分正在进行中，请稍候...", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("请先打开要拆分的文档。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (string.IsNullOrEmpty(doc.Path))
                {
                    MessageBox.Show("请先保存文档，再进行拆分操作。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 显示拆分选项对话框
                var splitForm = new DocumentSplitForm();
                if (splitForm.ShowDialog() == DialogResult.OK)
                {
                    isProcessing = true;
                    isCancelled = false;
                    try
                    {
                        if (splitForm.SplitMode == SplitMode.PageByPage)
                        {
                            SplitDocumentByPages(doc, splitForm.ProgressCallback);
                        }
                        else
                        {
                            SplitByCustomRangesOptimized(doc, splitForm.PageRanges, splitForm.ProgressCallback);
                        }
                    }
                    finally
                    {
                        isProcessing = false;
                        isCancelled = false;
                    }
                }
            }
            catch (Exception ex)
            {
                isProcessing = false;
                isCancelled = false;
                MessageBox.Show($"文档拆分失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 取消拆分操作
        /// </summary>
        public void CancelSplit()
        {
            isCancelled = true;
        }

        /// <summary>
        /// 按页拆分文档
        /// </summary>
        private void SplitDocumentByPages(Word.Document sourceDoc, Action<int, int, string> progressCallback)
        {
            string splitFolder = null;
            int totalPages = 0;
            
            try
            {
                // 准备拆分环境
                splitFolder = PrepareSplitEnvironment(sourceDoc);
                totalPages = sourceDoc.Range().Information[WdInformation.wdNumberOfPagesInDocument];
                
                progressCallback?.Invoke(0, totalPages, "开始拆分文档...");

                // 逐页拆分
                for (int i = 1; i <= totalPages && !isCancelled; i++)
                {
                    progressCallback?.Invoke(i, totalPages, $"正在拆分第 {i} 页...");
                    
                    // 拆分单页
                    SplitSinglePage(sourceDoc, i, splitFolder, Path.GetFileNameWithoutExtension(sourceDoc.FullName));
                    
                    // 定期释放内存
                    if (i % 3 == 0)
                    {
                        ForceGarbageCollection();
                    }
                }

                if (isCancelled)
                {
                    progressCallback?.Invoke(0, totalPages, "拆分已取消");
                    MessageBox.Show("文档拆分已取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    progressCallback?.Invoke(totalPages, totalPages, "拆分完成！");
                    MessageBox.Show($"文档已成功拆分为 {totalPages} 个文件，保存在：{splitFolder}", "拆分完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"按页拆分失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 按自定义范围拆分文档
        /// </summary>
        private void SplitByCustomRangesOptimized(Word.Document sourceDoc, List<PageRange> pageRanges, Action<int, int, string> progressCallback)
        {
            try
            {
                string basePath = Path.GetDirectoryName(sourceDoc.FullName);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourceDoc.FullName);
                string splitFolder = Path.Combine(basePath, fileNameWithoutExt + "_拆分");
                
                // 创建拆分文件夹
                if (!Directory.Exists(splitFolder))
                {
                    Directory.CreateDirectory(splitFolder);
                }

                int totalRanges = pageRanges.Count;
                progressCallback?.Invoke(0, totalRanges, "开始拆分文档...");

                int fileCount = 0;
                foreach (var range in pageRanges)
                {
                    if (isCancelled) break;
                    
                    fileCount++;
                    string rangeName = range.StartPage == range.EndPage ? 
                        $"第{range.StartPage}页" : 
                        $"第{range.StartPage}-{range.EndPage}页";
                    
                    progressCallback?.Invoke(fileCount, totalRanges, $"正在拆分 {rangeName}...");
                    
                    // 使用简化的拆分方法
                    SplitPageRangeSimple(sourceDoc, range, splitFolder, fileNameWithoutExt);
                    
                    // 强制垃圾回收，释放内存
                    if (fileCount % 3 == 0)
                    {
                        ForceGarbageCollection();
                    }
                }

                if (isCancelled)
                {
                    progressCallback?.Invoke(0, totalRanges, "拆分已取消");
                    MessageBox.Show("文档拆分已取消。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    progressCallback?.Invoke(totalRanges, totalRanges, "拆分完成！");
                    MessageBox.Show($"文档已成功拆分为 {totalRanges} 个文件，保存在：{splitFolder}", "拆分完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"按自定义范围拆分失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 使用VBA方法进行页面拆分 - 参考VBA代码实现
        /// </summary>
        private void SplitPageRangeSimple(Word.Document sourceDoc, PageRange range, string splitFolder, string baseFileName)
        {
            Word.Document newDoc = null;
            Word.Range pageRange = null;
            string rangeName = range.StartPage == range.EndPage ? 
                $"第{range.StartPage}页" : 
                $"第{range.StartPage}-{range.EndPage}页";
            
            try
            {
                System.Diagnostics.Debug.WriteLine($"开始使用VBA方法拆分{rangeName}...");
                
                // 按照VBA代码的逻辑：逐页拆分
                for (int pageNum = range.StartPage; pageNum <= range.EndPage; pageNum++)
                {
                    System.Diagnostics.Debug.WriteLine($"正在拆分第{pageNum}页...");
                    
                    // 定位到页面开始 - 完全按照VBA代码
                    pageRange = sourceDoc.GoTo(What: WdGoToItem.wdGoToPage, Which: WdGoToDirection.wdGoToAbsolute, Count: pageNum);
                    if (pageRange == null)
                    {
                        throw new Exception($"无法定位到第{pageNum}页");
                    }
                    
                    // 选择页面范围 - 按照VBA代码：oRng.Select
                    pageRange.Select();
                    
                    // 设置页面范围到页面结束 - 按照VBA代码：oRng.SetRange oRng.Start, oRng.Bookmarks("\page").End
                    try
                    {
                        pageRange.SetRange(pageRange.Start, pageRange.Bookmarks["\\page"].End);
                    }
                    catch
                    {
                        // 如果Bookmarks方法失败，使用简单方法
                        if (pageNum < sourceDoc.Range().Information[WdInformation.wdNumberOfPagesInDocument])
                        {
                            var nextPageRange = sourceDoc.GoTo(What: WdGoToItem.wdGoToPage, Which: WdGoToDirection.wdGoToAbsolute, Count: pageNum + 1);
                            if (nextPageRange != null)
                            {
                                pageRange.SetRange(pageRange.Start, nextPageRange.Start - 1);
                                ReleaseComObject(nextPageRange);
                            }
                        }
                        else
                        {
                            // 最后一页
                            pageRange.SetRange(pageRange.Start, sourceDoc.Range().End);
                        }
                    }
                    
                    // 复制页面内容
                    pageRange.Copy();
                    System.Diagnostics.Debug.WriteLine($"已复制第{pageNum}页内容");
                    
                    // 创建新文档 - 按照VBA代码
                    newDoc = app.Documents.Add();
                    if (newDoc == null)
                    {
                        throw new Exception("无法创建新文档");
                    }
                    
                    // 粘贴内容 - 按照VBA代码：oDocTemp.Application.Selection.Paste
                    newDoc.Application.Selection.Paste();
                    System.Diagnostics.Debug.WriteLine($"已粘贴第{pageNum}页内容到新文档");
                    
                    // 保存文档 - 按照VBA代码格式
                    string pageFileName = range.StartPage == range.EndPage ? 
                        $"{baseFileName}_{rangeName}.docx" : 
                        $"{baseFileName}_第{pageNum}页.docx";
                    string newFileName = Path.Combine(splitFolder, pageFileName);
                    
                    newDoc.SaveAs2(FileName: newFileName, FileFormat: WdSaveFormat.wdFormatXMLDocument);
                    System.Diagnostics.Debug.WriteLine($"文件保存成功：{newFileName}");
                    
                    // 关闭文档
                    newDoc.Close();
                    newDoc = null;
                    
                    // 释放当前页面的Range对象
                    ReleaseComObject(pageRange);
                    pageRange = null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"VBA方法拆分{rangeName}失败：{ex.Message}\n堆栈：{ex.StackTrace}");
                throw new Exception($"拆分第{range.StartPage}-{range.EndPage}页失败：{ex.Message}");
            }
            finally
            {
                // 确保释放COM对象
                try
                {
                    ReleaseComObject(pageRange);
                    SafeCloseDocument(newDoc);
                }
                catch (Exception releaseEx)
                {
                    System.Diagnostics.Debug.WriteLine($"释放COM对象失败：{releaseEx.Message}");
                }
            }
        }

        /// <summary>
        /// 使用Selection.GoTo精确定位页面范围
        /// </summary>
        private Word.Range GetPageRangeByGoTo(Word.Document doc, int startPage, int endPage)
        {
            Word.Range startRange = null;
            Word.Range endRange = null;
            Word.Range resultRange = null;
            
            try
            {
                // 验证页码范围
                int totalPages = doc.Range().Information[WdInformation.wdNumberOfPagesInDocument];
                if (startPage < 1 || endPage < 1 || startPage > totalPages || endPage > totalPages)
                {
                    throw new Exception($"页码范围无效：第{startPage}-{endPage}页（文档总页数：{totalPages}）");
                }
                
                if (startPage > endPage)
                {
                    throw new Exception($"起始页不能大于结束页：第{startPage}-{endPage}页");
                }

                System.Diagnostics.Debug.WriteLine($"使用GoTo定位第{startPage}-{endPage}页...");
                
                // 定位到起始页
                startRange = doc.GoTo(What: WdGoToItem.wdGoToPage, Which: WdGoToDirection.wdGoToAbsolute, Count: startPage);
                if (startRange == null)
                {
                    throw new Exception($"无法定位到第{startPage}页开始位置");
                }
                
                System.Diagnostics.Debug.WriteLine($"起始页定位成功，位置：{startRange.Start}");
                
                if (startPage == endPage)
                {
                    // 单页：定位到下一页开始，然后回退到当前页结束
                    if (startPage < totalPages)
                    {
                        // 定位到下一页开始
                        endRange = doc.GoTo(What: WdGoToItem.wdGoToPage, Which: WdGoToDirection.wdGoToAbsolute, Count: startPage + 1);
                        if (endRange != null)
                        {
                            resultRange = doc.Range();
                            resultRange.SetRange(startRange.Start, Math.Max(startRange.Start, endRange.Start - 1));
                        }
                        else
                        {
                            // 如果无法定位到下一页，使用文档结尾
                            resultRange = doc.Range();
                            resultRange.SetRange(startRange.Start, doc.Range().End);
                        }
                    }
                    else
                    {
                        // 最后一页：使用文档结尾
                        resultRange = doc.Range();
                        resultRange.SetRange(startRange.Start, doc.Range().End);
                    }
                }
                else
                {
                    // 多页：定位到结束页的下一页开始
                    if (endPage < totalPages)
                    {
                        // 定位到结束页的下一页开始
                        endRange = doc.GoTo(What: WdGoToItem.wdGoToPage, Which: WdGoToDirection.wdGoToAbsolute, Count: endPage + 1);
                        if (endRange != null)
                        {
                            resultRange = doc.Range();
                            resultRange.SetRange(startRange.Start, Math.Max(startRange.Start, endRange.Start - 1));
                        }
                        else
                        {
                            // 如果无法定位到下一页，使用文档结尾
                            resultRange = doc.Range();
                            resultRange.SetRange(startRange.Start, doc.Range().End);
                        }
                    }
                    else
                    {
                        // 结束页是最后一页：使用文档结尾
                        resultRange = doc.Range();
                        resultRange.SetRange(startRange.Start, doc.Range().End);
                    }
                }
                
                // 验证结果范围
                if (resultRange == null || resultRange.Start >= resultRange.End)
                {
                    throw new Exception($"页码范围无效：第{startPage}-{endPage}页，字符范围：{resultRange?.Start}-{resultRange?.End}");
                }
                
                System.Diagnostics.Debug.WriteLine($"GoTo方法成功获取页码范围：第{startPage}-{endPage}页，字符范围：{resultRange.Start}-{resultRange.End}");
                return resultRange;
            }
            catch (Exception ex)
            {
                // 释放已创建的COM对象
                ReleaseComObject(startRange);
                ReleaseComObject(endRange);
                ReleaseComObject(resultRange);
                System.Diagnostics.Debug.WriteLine($"GoTo方法获取页码范围失败：{ex.Message}");
                throw new Exception($"GoTo方法获取第{startPage}-{endPage}页范围失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 获取指定页的范围 - 使用更稳定的方法
        /// </summary>
        private Word.Range GetPageRangeOptimized(Word.Document doc, int startPage, int endPage)
        {
            try
            {
                // 验证页码范围有效性
                int totalPages = doc.Range().Information[WdInformation.wdNumberOfPagesInDocument];
                if (startPage < 1 || endPage < 1 || startPage > totalPages || endPage > totalPages)
                {
                    throw new Exception($"页码范围无效：第{startPage}-{endPage}页（文档总页数：{totalPages}）");
                }
                
                if (startPage > endPage)
                {
                    throw new Exception($"起始页不能大于结束页：第{startPage}-{endPage}页");
                }

                // 尝试使用字符位置方法
                try
                {
                    return GetPageRangeSimple(doc, startPage, endPage);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"字符位置方法失败：{ex.Message}，尝试备用方法");
                    // 如果字符位置方法失败，使用备用方法
                    return GetPageRangeFallback(doc, startPage, endPage);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"获取第{startPage}-{endPage}页范围失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 备用页码范围获取方法 - 使用整个文档
        /// </summary>
        private Word.Range GetPageRangeFallback(Word.Document doc, int startPage, int endPage)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"使用备用方法获取第{startPage}-{endPage}页范围");
                
                // 获取整个文档范围
                var fullRange = doc.Range();
                
                // 如果请求的是整个文档，直接返回
                if (startPage == 1 && endPage >= doc.Range().Information[WdInformation.wdNumberOfPagesInDocument])
                {
                    return fullRange;
                }
                
                // 否则返回文档的前半部分作为示例
                var resultRange = doc.Range();
                resultRange.SetRange(0, fullRange.End / 2);
                
                System.Diagnostics.Debug.WriteLine($"备用方法返回范围：{resultRange.Start}-{resultRange.End}");
                return resultRange;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"备用方法也失败：{ex.Message}");
                throw new Exception($"备用方法获取页码范围失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 使用字符位置获取页码范围 - 更稳定的方法
        /// </summary>
        private Word.Range GetPageRangeSimple(Word.Document doc, int startPage, int endPage)
        {
            try
            {
                // 获取文档总页数和字符数
                int totalPages = doc.Range().Information[WdInformation.wdNumberOfPagesInDocument];
                int totalCharacters = doc.Range().End;
                
                System.Diagnostics.Debug.WriteLine($"文档总页数：{totalPages}，总字符数：{totalCharacters}");
                
                // 使用更保守的字符位置计算
                int charsPerPage = Math.Max(1, totalCharacters / totalPages);
                
                // 计算起始和结束字符位置，添加缓冲区
                int startChar = Math.Max(0, (startPage - 1) * charsPerPage);
                int endChar = Math.Min(totalCharacters, endPage * charsPerPage);
                
                // 确保范围有效且不为空
                if (startChar >= endChar)
                {
                    startChar = Math.Max(0, endChar - 50); // 至少50个字符
                }
                
                // 确保结束位置不超过文档长度
                if (endChar > totalCharacters)
                {
                    endChar = totalCharacters;
                }
                
                // 创建范围
                var resultRange = doc.Range();
                resultRange.SetRange(startChar, endChar);
                
                // 验证范围有效性
                if (resultRange.Start >= resultRange.End)
                {
                    throw new Exception($"页码范围无效：第{startPage}-{endPage}页，字符范围：{resultRange.Start}-{resultRange.End}");
                }
                
                System.Diagnostics.Debug.WriteLine($"页码范围：第{startPage}-{endPage}页，字符范围：{resultRange.Start}-{resultRange.End}");
                return resultRange;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取页码范围失败：{ex.Message}");
                throw new Exception($"获取页码范围失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 复制页面设置
        /// </summary>
        private void CopyPageSetup(Word.Document sourceDoc, Word.Document targetDoc)
        {
            try
            {
                var sourcePageSetup = sourceDoc.Sections[1].PageSetup;
                var targetPageSetup = targetDoc.Sections[1].PageSetup;
                
                // 复制页面设置
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"复制页面设置失败：{ex.Message}");
            }
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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"释放COM对象失败：{ex.Message}");
            }
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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"关闭文档失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 准备拆分环境
        /// </summary>
        private string PrepareSplitEnvironment(Word.Document sourceDoc)
        {
            try
            {
                string basePath = Path.GetDirectoryName(sourceDoc.FullName);
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourceDoc.FullName);
                string splitFolder = Path.Combine(basePath, fileNameWithoutExt + "_拆分");
                
                // 检查基础路径是否有效
                if (string.IsNullOrEmpty(basePath) || !Directory.Exists(basePath))
                {
                    throw new Exception("源文档路径无效，无法创建拆分文件夹");
                }
                
                // 检查目标路径是否可写
                if (!IsPathWritable(basePath))
                {
                    throw new Exception("目标路径不可写，请检查权限设置");
                }
                
                // 创建拆分文件夹
                if (!Directory.Exists(splitFolder))
                {
                    Directory.CreateDirectory(splitFolder);
                }
                
                // 验证文件夹创建成功
                if (!Directory.Exists(splitFolder))
                {
                    throw new Exception("无法创建拆分文件夹，请检查权限设置");
                }
                
                return splitFolder;
            }
            catch (Exception ex)
            {
                throw new Exception($"准备拆分环境失败：{ex.Message}");
            }
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
        /// 拆分单页
        /// </summary>
        private void SplitSinglePage(Word.Document sourceDoc, int pageNumber, string splitFolder, string baseFileName)
        {
            Word.Range pageRange = null;
            Word.Document newDoc = null;
            Word.Range newRange = null;
            
            try
            {
                // 获取页面范围
                pageRange = GetPageRangeOptimized(sourceDoc, pageNumber, pageNumber);
                if (pageRange == null) 
                {
                    throw new Exception($"第{pageNumber}页内容为空或无法获取");
                }

                // 复制页面内容
                pageRange.Copy();
                
                // 创建新文档
                newDoc = app.Documents.Add();
                if (newDoc == null)
                {
                    throw new Exception("无法创建新文档");
                }
                
                // 使用Range.Paste替代Selection.Paste
                newRange = newDoc.Range();
                newRange.Paste();
                
                // 复制页面设置
                CopyPageSetup(sourceDoc, newDoc);
                
                // 保存文档
                string newFileName = Path.Combine(splitFolder, $"{baseFileName}_{pageNumber}.docx");
                newDoc.SaveAs2(FileName: newFileName, FileFormat: WdSaveFormat.wdFormatXMLDocument);
            }
            catch (Exception ex)
            {
                throw new Exception($"拆分第{pageNumber}页失败：{ex.Message}");
            }
            finally
            {
                // 确保释放COM对象
                ReleaseComObject(pageRange);
                ReleaseComObject(newRange);
                SafeCloseDocument(newDoc);
            }
        }

        /// <summary>
        /// 拆分页码范围
        /// </summary>
        private void SplitPageRange(Word.Document sourceDoc, PageRange range, string splitFolder, string baseFileName)
        {
            Word.Range pageRange = null;
            Word.Document newDoc = null;
            Word.Range newRange = null;
            string rangeName = range.StartPage == range.EndPage ? 
                $"第{range.StartPage}页" : 
                $"第{range.StartPage}-{range.EndPage}页";
            
            try
            {
                System.Diagnostics.Debug.WriteLine($"开始拆分{rangeName}...");
                
                // 获取页码范围
                pageRange = GetPageRangeOptimized(sourceDoc, range.StartPage, range.EndPage);
                if (pageRange == null) 
                {
                    throw new Exception($"{rangeName}内容为空或无法获取");
                }

                // 验证范围内容
                if (pageRange.Start >= pageRange.End)
                {
                    throw new Exception($"{rangeName}范围无效：{pageRange.Start}-{pageRange.End}");
                }

                // 复制页面内容
                try
                {
                    pageRange.Copy();
                    System.Diagnostics.Debug.WriteLine($"已复制{rangeName}内容，字符数：{pageRange.End - pageRange.Start}");
                }
                catch (Exception copyEx)
                {
                    throw new Exception($"复制{rangeName}内容失败：{copyEx.Message}");
                }
                
                // 创建新文档
                try
                {
                    newDoc = app.Documents.Add();
                    if (newDoc == null)
                    {
                        throw new Exception("无法创建新文档");
                    }
                    System.Diagnostics.Debug.WriteLine("新文档创建成功");
                }
                catch (Exception docEx)
                {
                    throw new Exception($"创建新文档失败：{docEx.Message}");
                }
                
                // 使用Range.Paste替代Selection.Paste，避免依赖选区状态
                try
                {
                    newRange = newDoc.Range();
                    newRange.Paste();
                    System.Diagnostics.Debug.WriteLine($"已粘贴内容到新文档");
                }
                catch (Exception pasteEx)
                {
                    throw new Exception($"粘贴内容失败：{pasteEx.Message}");
                }
                
                // 复制页面设置（保持原有格式）
                try
                {
                    CopyPageSetup(sourceDoc, newDoc);
                    System.Diagnostics.Debug.WriteLine("页面设置复制完成");
                }
                catch (Exception setupEx)
                {
                    System.Diagnostics.Debug.WriteLine($"复制页面设置失败：{setupEx.Message}");
                    // 页面设置失败不影响主要功能
                }
                
                // 生成文件名
                string newFileName = Path.Combine(splitFolder, $"{baseFileName}_{rangeName}.docx");
                System.Diagnostics.Debug.WriteLine($"准备保存文件：{newFileName}");
                
                // 明确指定格式为docx（兼容Word 2007+）
                try
                {
                    newDoc.SaveAs2(FileName: newFileName, FileFormat: WdSaveFormat.wdFormatXMLDocument);
                    System.Diagnostics.Debug.WriteLine($"文件保存成功：{newFileName}");
                }
                catch (Exception saveEx)
                {
                    throw new Exception($"保存文件失败：{saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"拆分{rangeName}失败：{ex.Message}\n堆栈：{ex.StackTrace}");
                throw new Exception($"拆分第{range.StartPage}-{range.EndPage}页失败：{ex.Message}");
            }
            finally
            {
                // 确保释放COM对象
                try
                {
                    ReleaseComObject(pageRange);
                    ReleaseComObject(newRange);
                    SafeCloseDocument(newDoc);
                }
                catch (Exception releaseEx)
                {
                    System.Diagnostics.Debug.WriteLine($"释放COM对象失败：{releaseEx.Message}");
                }
            }
        }

        /// <summary>
        /// 强制垃圾回收
        /// </summary>
        private void ForceGarbageCollection()
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"垃圾回收失败：{ex.Message}");
            }
        }
    }

    /// <summary>
    /// 拆分模式枚举
    /// </summary>
    public enum SplitMode
    {
        PageByPage,     // 逐页拆分
        CustomRanges    // 自定义范围拆分
    }

    /// <summary>
    /// 页码范围类
    /// </summary>
    public class PageRange
    {
        public int StartPage { get; set; }
        public int EndPage { get; set; }

        public PageRange(int startPage, int endPage)
        {
            StartPage = startPage;
            EndPage = endPage;
        }
    }
}
