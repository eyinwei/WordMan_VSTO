using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordMan.SplitAndMerge
{
    public class DocumentMerger
    {
        private readonly Microsoft.Office.Interop.Word.Application app;

        public DocumentMerger(Microsoft.Office.Interop.Word.Application application)
        {
            app = application;
        }

        private void ReleaseComObject(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }

        private void CopyPageSetup(Microsoft.Office.Interop.Word.Document sourceDoc, Microsoft.Office.Interop.Word.Document targetDoc, int targetSectionIndex = 1)
        {
            if (sourceDoc?.Sections == null || targetDoc?.Sections == null) return;
            if (targetSectionIndex < 1 || targetSectionIndex > targetDoc.Sections.Count) return;

            var sourcePageSetup = sourceDoc.Sections[1].PageSetup;
            var targetPageSetup = targetDoc.Sections[targetSectionIndex].PageSetup;
            
            targetPageSetup.Orientation = sourcePageSetup.Orientation;
            targetPageSetup.PageWidth = sourcePageSetup.PageWidth;
            targetPageSetup.PageHeight = sourcePageSetup.PageHeight;
            targetPageSetup.LeftMargin = sourcePageSetup.LeftMargin;
            targetPageSetup.RightMargin = sourcePageSetup.RightMargin;
            targetPageSetup.TopMargin = sourcePageSetup.TopMargin;
            targetPageSetup.BottomMargin = sourcePageSetup.BottomMargin;
        }

        private void CopyHeaderFooterToSection(Microsoft.Office.Interop.Word.Document sourceDoc, Microsoft.Office.Interop.Word.Document targetDoc, int targetSectionIndex)
        {
            if (sourceDoc?.Sections == null || targetDoc?.Sections == null) return;
            if (targetSectionIndex < 1 || targetSectionIndex > targetDoc.Sections.Count) return;

            var sourceSection = sourceDoc.Sections[1];
            var targetSection = targetDoc.Sections[targetSectionIndex];

            // 复制奇偶页页眉页脚设置
            targetSection.PageSetup.DifferentFirstPageHeaderFooter = sourceSection.PageSetup.DifferentFirstPageHeaderFooter;
            targetSection.PageSetup.OddAndEvenPagesHeaderFooter = sourceSection.PageSetup.OddAndEvenPagesHeaderFooter;
        }

        private void ReplaceHeaderFooterAfterInsert(Microsoft.Office.Interop.Word.Document sourceDoc, Microsoft.Office.Interop.Word.Document targetDoc, int targetSectionIndex)
        {
            if (sourceDoc?.Sections == null || targetDoc?.Sections == null) return;
            if (targetSectionIndex < 1 || targetSectionIndex > targetDoc.Sections.Count) return;

            var sourceSection = sourceDoc.Sections[1];
            var targetSection = targetDoc.Sections[targetSectionIndex];

            // 获取源文档的页眉页脚内容
            var sourceHeaderText = sourceSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;
            var sourceFooterText = sourceSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;

            // 如果源文档有页眉页脚内容，则替换目标文档的页眉页脚
            if (!string.IsNullOrEmpty(sourceHeaderText.Trim()))
            {
                targetSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = sourceHeaderText;
            }

            if (!string.IsNullOrEmpty(sourceFooterText.Trim()))
            {
                targetSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = sourceFooterText;
            }
        }

        private void ProcessDocument(string filePath, Microsoft.Office.Interop.Word.Document mergedDoc, int index, MergeOptions options)
        {
            Microsoft.Office.Interop.Word.Document sourceDoc = null;
            
            try
            {
                sourceDoc = app.Documents.Open(filePath, ReadOnly: true, AddToRecentFiles: false);
                sourceDoc.Windows[1].Visible = false;
                
                var targetRange = mergedDoc.Range();
                targetRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                
                if (index > 0 && options.AddPageBreaks)
                {
                    if (options.UseSectionBreak)
                    {
                        targetRange.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                        targetRange = mergedDoc.Range();
                        targetRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                        
                        CopyPageSetup(sourceDoc, mergedDoc, mergedDoc.Sections.Count);
                        CopyHeaderFooterToSection(sourceDoc, mergedDoc, mergedDoc.Sections.Count);
                    }
                    else
                    {
                        targetRange.InsertBreak(WdBreakType.wdPageBreak);
                        targetRange = mergedDoc.Range();
                        targetRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    }
                }
                else if (index == 0)
                {
                    CopyPageSetup(sourceDoc, mergedDoc);
                }
                
                targetRange.InsertFile(filePath);
                
                if (index > 0 && options.AddPageBreaks && options.UseSectionBreak)
                {
                    mergedDoc.Range().Collapse(WdCollapseDirection.wdCollapseEnd);
                    ReplaceHeaderFooterAfterInsert(sourceDoc, mergedDoc, mergedDoc.Sections.Count);
                }
            }
            finally
            {
                if (sourceDoc != null)
                {
                    sourceDoc.Close(SaveChanges: false);
                    ReleaseComObject(sourceDoc);
                }
            }
        }

        public void ShowMergeDialog()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "选择要合并的Word文档",
                Filter = "Word文档 (*.docx)|*.docx|Word文档 (*.doc)|*.doc|所有文件 (*.*)|*.*",
                Multiselect = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filePaths = ValidateFiles(openFileDialog.FileNames.ToList());
                if (filePaths.Count < 2)
                {
                    MessageBox.Show("请至少选择2个有效文档进行合并。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var mergeForm = new DocumentMergeForm(filePaths);
                if (mergeForm.ShowDialog() == DialogResult.OK)
                {
                    MergeDocuments(mergeForm.SelectedFiles, mergeForm.MergeOptions);
                }
            }
        }

        private List<string> ValidateFiles(List<string> filePaths)
        {
            var validFiles = new List<string>();
            
            foreach (var filePath in filePaths)
            {
                if (File.Exists(filePath) && IsValidWordDocument(filePath))
                {
                    validFiles.Add(filePath);
                }
            }
            
            return validFiles;
        }

        private bool IsValidWordDocument(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            return extension == ".docx" || extension == ".doc";
        }

        private void MergeDocuments(List<string> filePaths, MergeOptions options)
        {
            var originalScreenUpdating = app.ScreenUpdating;
            var originalDisplayAlerts = app.DisplayAlerts;
            Microsoft.Office.Interop.Word.Document mergedDoc = null;
            
            try
            {
                mergedDoc = app.Documents.Add();
                mergedDoc.Windows[1].Visible = false;
                
                app.ScreenUpdating = false;
                app.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                for (int i = 0; i < filePaths.Count; i++)
                {
                    ProcessDocument(filePaths[i], mergedDoc, i, options);
                }

                app.ScreenUpdating = originalScreenUpdating;
                app.DisplayAlerts = originalDisplayAlerts;
                
                mergedDoc.Windows[1].Visible = true;
                mergedDoc.Activate();
                
                MessageBox.Show("文档合并完成！", "合并完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = WdAlertLevel.wdAlertsAll;
                
                if (mergedDoc != null)
                {
                    mergedDoc.Close(SaveChanges: false);
                    ReleaseComObject(mergedDoc);
                }
                
                MessageBox.Show($"合并失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
