using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan_VSTO.SplitAndMerge
{
    public class DocumentMerger
    {
        private Word.Application app;

        public DocumentMerger(Word.Application application)
        {
            app = application;
        }

        /// <summary>
        /// 安全释放COM对象
        /// </summary>
        private void ReleaseComObject(object comObject)
        {
            try
            {
                if (comObject != null && Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"释放COM对象失败：{ex.Message}");
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

        /// <summary>
        /// 复制页面设置从源文档到目标文档
        /// </summary>
        private void CopyPageSetup(Word.Document sourceDoc, Word.Document targetDoc, int targetSectionIndex = 1)
        {
            try
            {
                if (sourceDoc?.Sections == null || targetDoc?.Sections == null)
                    return;

                var sourcePageSetup = sourceDoc.Sections[1].PageSetup;
                var targetPageSetup = targetDoc.Sections[targetSectionIndex].PageSetup;
                
                // 复制基本页面设置
                targetPageSetup.Orientation = sourcePageSetup.Orientation;
                targetPageSetup.PageWidth = sourcePageSetup.PageWidth;
                targetPageSetup.PageHeight = sourcePageSetup.PageHeight;
                targetPageSetup.LeftMargin = sourcePageSetup.LeftMargin;
                targetPageSetup.RightMargin = sourcePageSetup.RightMargin;
                targetPageSetup.TopMargin = sourcePageSetup.TopMargin;
                targetPageSetup.BottomMargin = sourcePageSetup.BottomMargin;
                
                // 复制栏设置
                if (sourcePageSetup.TextColumns.Count != targetPageSetup.TextColumns.Count)
                {
                    targetPageSetup.TextColumns.SetCount(sourcePageSetup.TextColumns.Count);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"复制页面设置失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 显示文档合并对话框
        /// </summary>
        public void ShowMergeDialog()
        {
            try
            {
                // 显示文件选择对话框
                var openFileDialog = new OpenFileDialog
                {
                    Title = "选择要合并的Word文档",
                    Filter = "Word文档 (*.docx)|*.docx|Word文档 (*.doc)|*.doc|所有文件 (*.*)|*.*",
                    Multiselect = true
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var filePaths = openFileDialog.FileNames.ToList();
                    if (filePaths.Count < 2)
                    {
                        MessageBox.Show("请至少选择2个文档进行合并。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 显示合并选项对话框
                    var mergeForm = new DocumentMergeForm(filePaths);
                    if (mergeForm.ShowDialog() == DialogResult.OK)
                    {
                        // 直接执行合并（已优化为后台模式）
                        MergeDocuments(mergeForm.SelectedFiles, mergeForm.MergeOptions);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文档合并失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 合并文档
        /// </summary>
        private void MergeDocuments(List<string> filePaths, MergeOptions options)
        {
            // 保存原始状态
            var originalScreenUpdating = app.ScreenUpdating;
            var originalDisplayAlerts = app.DisplayAlerts;
            var originalDoc = app.ActiveDocument;
            var originalWindow = app.ActiveWindow;
            Word.Document mergedDoc = null;
            Word.Document firstDoc = null;
            
            try
            {
                // 创建新文档（后台创建，不显示）
                mergedDoc = app.Documents.Add();
                mergedDoc.Windows[1].Visible = false; // 隐藏新文档窗口
                
                // 设置第一个文档的页面设置，避免影响后续文档
                if (filePaths.Count > 0)
                {
                    try
                    {
                        firstDoc = app.Documents.Open(filePaths[0], ReadOnly: true, AddToRecentFiles: false);
                        
                        // 立即隐藏第一个文档窗口
                        if (firstDoc?.Windows.Count > 0)
                        {
                            firstDoc.Windows[1].Visible = false;
                        }
                        
                        // 使用统一的页面设置复制方法
                        CopyPageSetup(firstDoc, mergedDoc);
                        
                        firstDoc.Close(SaveChanges: false);
                        ReleaseComObject(firstDoc);
                        firstDoc = null;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"设置第一个文档页面设置失败：{ex.Message}");
                    }
                }
                
                // 立即恢复原文档激活状态，防止被关闭
                if (originalDoc != null && originalDoc != mergedDoc)
                {
                    try
                    {
                        originalDoc.Activate();
                        if (originalWindow != null)
                        {
                            originalWindow.Activate();
                        }
                    }
                    catch
                    {
                        // 如果激活失败，尝试重新获取活动文档
                        try
                        {
                            originalDoc = app.ActiveDocument;
                            originalWindow = app.ActiveWindow;
                        }
                        catch { }
                    }
                }
                
                // 保持ScreenUpdating关闭状态，提升性能
                app.ScreenUpdating = false;
                app.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                // 逐个合并文档到新文档（优化版本）
                
                for (int i = 0; i < filePaths.Count; i++)
                {
                    var filePath = filePaths[i];
                    Word.Document tempDoc = null;
                    
                    try
                    {
                        // 检查合并文档是否仍然有效
                        if (mergedDoc == null || mergedDoc.Sections == null)
                        {
                            throw new Exception("合并文档对象已失效");
                        }
                        
                        // 确保合并文档窗口保持隐藏
                        if (mergedDoc.Windows.Count > 0)
                        {
                            mergedDoc.Windows[1].Visible = false;
                        }
                        
                        // 在合并文档中定位到末尾
                        var targetRange = mergedDoc.Range();
                        targetRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                        
                        // 如果不是第一个文档，先插入分页符或分节符
                        if (i > 0 && options.AddPageBreaks)
                        {
                            if (options.UseSectionBreak)
                            {
                                // 插入分节符（下一页）
                                targetRange.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                            }
                            else
                            {
                                // 插入分页符
                                targetRange.InsertBreak(WdBreakType.wdPageBreak);
                            }
                            
                            // 重新定位到新节的末尾
                            targetRange = mergedDoc.Range();
                            targetRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                            
                            // 如果使用分节符，获取源文档的页面设置并应用到新节
                            if (options.UseSectionBreak)
                            {
                                try
                                {
                                    // 检查文件是否存在
                                    if (!System.IO.File.Exists(filePath))
                                    {
                                        throw new Exception($"文件不存在：{filePath}");
                                    }
                                    
                                    tempDoc = app.Documents.Open(filePath, ReadOnly: true, AddToRecentFiles: false);
                                    
                                    // 立即隐藏临时文档窗口
                                    if (tempDoc?.Windows.Count > 0)
                                    {
                                        tempDoc.Windows[1].Visible = false;
                                    }
                                    
                                    // 使用统一的页面设置复制方法
                                    CopyPageSetup(tempDoc, mergedDoc, mergedDoc.Sections.Count);
                                    
                                    // 重新定位到新节的末尾（页面设置可能改变位置）
                                    targetRange = mergedDoc.Range();
                                    targetRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"复制页面设置失败：{ex.Message}");
                                    // 继续执行，不中断合并过程
                                }
                                finally
                                {
                                    // 确保临时文档被关闭并释放COM对象
                                    if (tempDoc != null)
                                    {
                                        try
                                        {
                                            tempDoc.Close(SaveChanges: false);
                                        }
                                        catch { }
                                        ReleaseComObject(tempDoc);
                                        tempDoc = null;
                                    }
                                }
                            }
                        }
                        
                        // 使用InsertFile方法直接插入文档内容（最快的方式）
                        targetRange.InsertFile(filePath);
                        
                        // 确保合并文档窗口保持隐藏
                        if (mergedDoc.Windows.Count > 0)
                        {
                            mergedDoc.Windows[1].Visible = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        // 确保临时文档被关闭并释放COM对象
                        if (tempDoc != null)
                        {
                            try
                            {
                                tempDoc.Close(SaveChanges: false);
                            }
                            catch { }
                            ReleaseComObject(tempDoc);
                            tempDoc = null;
                        }
                        
                        MessageBox.Show($"合并文件 {Path.GetFileName(filePath)} 时出错：{ex.Message}", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                // 恢复Word应用程序设置
                app.ScreenUpdating = originalScreenUpdating;
                app.DisplayAlerts = originalDisplayAlerts;
                
                // 显示合并后的文档
                mergedDoc.Windows[1].Visible = true; // 显示新文档窗口
                mergedDoc.Activate();
                
                // 释放COM对象
                ReleaseComObject(mergedDoc);
                mergedDoc = null;
                
                // 强制垃圾回收
                ForceGarbageCollection();
                
                MessageBox.Show($"文档合并完成！", "合并完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                // 确保在异常情况下也恢复Word设置和释放COM对象
                try
                {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = WdAlertLevel.wdAlertsAll;
                    
                    // 释放所有COM对象
                    if (mergedDoc != null)
                    {
                        try
                        {
                            mergedDoc.Close(SaveChanges: false);
                        }
                        catch { }
                        ReleaseComObject(mergedDoc);
                        mergedDoc = null;
                    }
                    
                    if (firstDoc != null)
                    {
                        try
                        {
                            firstDoc.Close(SaveChanges: false);
                        }
                        catch { }
                        ReleaseComObject(firstDoc);
                        firstDoc = null;
                    }
                    
                    // 强制垃圾回收
                    ForceGarbageCollection();
                    
                    // 确保原文档保持激活状态
                    if (originalDoc != null)
                    {
                        try
                        {
                            originalDoc.Activate();
                        }
                        catch { }
                    }
                }
                catch { }
                
                MessageBox.Show($"合并失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 复制样式（当前为简化实现）
        /// </summary>
        private void CopyStyles(Word.Document sourceDoc, Word.Document targetDoc)
        {
            try
            {
                // 由于Word API的限制，样式复制比较复杂
                // 这里只做简单处理，实际项目中可能需要更复杂的实现
                foreach (Word.Style sourceStyle in sourceDoc.Styles)
                {
                    try
                    {
                        // 检查目标文档是否已有同名样式
                        bool styleExists = false;
                        foreach (Word.Style targetStyle in targetDoc.Styles)
                        {
                            if (targetStyle.NameLocal == sourceStyle.NameLocal)
                            {
                                styleExists = true;
                                break;
                            }
                        }

                        // 如果样式不存在且不是内置样式，则复制
                        if (!styleExists && sourceStyle.BuiltIn == false)
                        {
                            // 这里可以添加样式复制的具体实现
                            // 由于Word API的限制，样式复制比较复杂，这里只做简单处理
                        }
                    }
                    catch
                    {
                        // 忽略单个样式复制失败
                    }
                }
            }
            catch (Exception ex)
            {
                // 样式复制失败不影响主要功能
                System.Diagnostics.Debug.WriteLine($"复制样式失败：{ex.Message}");
            }
        }


    }

    /// <summary>
    /// 合并选项类
    /// </summary>
    public class MergeOptions
    {
        public bool AddPageBreaks { get; set; } = true;
        public bool PreserveFormatting { get; set; } = true;
        public bool CopyStyles { get; set; } = true; // 默认复制样式
        public bool UseSectionBreak { get; set; } = true; // 默认使用分节符
    }


}
