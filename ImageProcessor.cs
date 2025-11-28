using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;

namespace WordMan
{
    public class ImageProcessor
    {
        #region 宽度刷相关字段
        private float copiedWidth = 0f;
        private Timer widthBrushTimer;
        private object lastSelectionStart = null;
        private object lastSelectionEnd = null;
        private DateTime lastActivityTime = DateTime.Now;
        private int lastSelectionHash = 0;
        #endregion

        #region 高度刷相关字段
        private float copiedHeight = 0f;
        private Timer heightBrushTimer;
        private object lastHeightSelectionStart = null;
        private object lastHeightSelectionEnd = null;
        private DateTime lastHeightActivityTime = DateTime.Now;
        private int lastHeightSelectionHash = 0;
        #endregion

        #region 宽度刷相关方法
        public void WidthBrush_Click(object sender, RibbonControlEventArgs e, Microsoft.Office.Tools.Ribbon.RibbonToggleButton 宽度刷)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (toggleButton.Checked)
                    ActivateWidthBrush(app, toggleButton);
                else
                    DeactivateWidthBrush(app);
            }
            catch
            {
                toggleButton.Checked = false;
                DeactivateWidthBrush(Globals.ThisAddIn.Application);
            }
        }

        private void ActivateWidthBrush(Word.Application app, Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton)
        {
            var selection = app.Selection;
            if (selection.InlineShapes.Count == 0 && selection.ShapeRange.Count == 0)
            {
                toggleButton.Checked = false;
                return;
            }

            float width = 0f;
            if (selection.InlineShapes.Count > 0)
                width = selection.InlineShapes[1].Width;
            else if (selection.ShapeRange.Count > 0)
                width = selection.ShapeRange[1].Width;

            if (width <= 0)
            {
                toggleButton.Checked = false;
                return;
            }

            copiedWidth = width;
            lastActivityTime = DateTime.Now;
            lastSelectionHash = GetSelectionHash(selection);
            StartWidthBrushMonitoring(app);
            app.StatusBar = $"宽度刷已激活 - 宽度: {width}磅 - 按ESC或点击空白区域退出";
        }

        public void DeactivateWidthBrush(Word.Application app)
        {
            widthBrushTimer?.Stop();
            widthBrushTimer?.Dispose();
            widthBrushTimer = null;
            lastSelectionStart = null;
            lastSelectionEnd = null;
            lastSelectionHash = 0;
            try { app.StatusBar = ""; } catch { }
        }

        private void StartWidthBrushMonitoring(Word.Application app)
        {
            widthBrushTimer = new Timer();
            widthBrushTimer.Interval = 100;
            widthBrushTimer.Tick += (s, timerE) =>
            {
                // 注意：这里需要传入宽度刷按钮的引用
                try
                {
                    var currentSelection = app.Selection;
                    int currentHash = GetSelectionHash(currentSelection);

                    // 检查退出条件
                    if (currentHash != lastSelectionHash)
                    {
                        bool isEmptyNow = currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                            (currentSelection.Type == Word.WdSelectionType.wdSelectionIP || currentSelection.Type == Word.WdSelectionType.wdSelectionNormal) &&
                            string.IsNullOrWhiteSpace(currentSelection.Text?.Replace("\r", ""));

                        TimeSpan timeSinceLastActivity = DateTime.Now - lastActivityTime;
                        if (isEmptyNow && timeSinceLastActivity.TotalMilliseconds < 500) { DeactivateWidthBrush(app); return; }

                        if (currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                            currentSelection.Paragraphs.Count > 0 &&
                            string.IsNullOrWhiteSpace(currentSelection.Paragraphs[1].Range.Text.Replace("\r", "")))
                        { DeactivateWidthBrush(app); return; }
                    }

                    // 长时间停留在空白区域退出
                    if (currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                        string.IsNullOrWhiteSpace(currentSelection.Text?.Replace("\r", "")) &&
                        (DateTime.Now - lastActivityTime).TotalMilliseconds > 1000)
                    { DeactivateWidthBrush(app); return; }

                    // 应用宽度
                    bool hasImageSelection = currentSelection.InlineShapes.Count > 0 || currentSelection.ShapeRange.Count > 0;
                    if (hasImageSelection)
                    {
                        bool isNewSelection = false;
                        try
                        {
                            if (lastSelectionStart == null || lastSelectionEnd == null ||
                                !lastSelectionStart.Equals(currentSelection.Start) || !lastSelectionEnd.Equals(currentSelection.End))
                            {
                                isNewSelection = true;
                                lastSelectionStart = currentSelection.Start;
                                lastSelectionEnd = currentSelection.End;
                                lastActivityTime = DateTime.Now;
                                lastSelectionHash = currentHash;
                            }
                        }
                        catch { isNewSelection = true; lastActivityTime = DateTime.Now; lastSelectionHash = currentHash; }

                        if (isNewSelection && ApplyWidthToSelection(currentSelection))
                            app.StatusBar = $"宽度刷: 已应用宽度 {copiedWidth}磅 - 按ESC或点击空白区域退出";
                    }
                    else if (currentHash != lastSelectionHash)
                    {
                        lastActivityTime = DateTime.Now;
                        lastSelectionHash = currentHash;
                    }
                }
                catch { }
            };
            widthBrushTimer.Start();
        }

        private int GetSelectionHash(Word.Selection selection)
        {
            try
            {
                int hash = 0;
                hash ^= selection.Start.GetHashCode();
                hash ^= selection.End.GetHashCode();
                hash ^= selection.Type.GetHashCode();
                hash ^= selection.InlineShapes.Count.GetHashCode();
                hash ^= selection.ShapeRange.Count.GetHashCode();
                if (!string.IsNullOrEmpty(selection.Text))
                    hash ^= selection.Text.GetHashCode();
                return hash;
            }
            catch { return DateTime.Now.Millisecond; }
        }

        private bool ApplyWidthToSelection(Word.Selection selection)
        {
            bool applied = false;
            try
            {
                if (selection.InlineShapes.Count > 0)
                {
                    foreach (Word.InlineShape shape in selection.InlineShapes)
                    {
                        if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeChart ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeSmartArt)
                        {
                            shape.Width = copiedWidth;
                            applied = true;
                        }
                    }
                }

                if (selection.ShapeRange.Count > 0)
                {
                    foreach (Word.Shape shape in selection.ShapeRange)
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoChart ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            shape.Width = copiedWidth;
                            applied = true;
                        }
                    }
                }
            }
            catch { }
            return applied;
        }
        #endregion

        #region 高度刷相关方法
        public void HeightBrush_Click(object sender, RibbonControlEventArgs e, Microsoft.Office.Tools.Ribbon.RibbonToggleButton 高度刷)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (toggleButton.Checked)
                    ActivateHeightBrush(app, toggleButton);
                else
                    DeactivateHeightBrush(app);
            }
            catch
            {
                toggleButton.Checked = false;
                DeactivateHeightBrush(Globals.ThisAddIn.Application);
            }
        }

        private void ActivateHeightBrush(Word.Application app, Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton)
        {
            var selection = app.Selection;
            if (selection.InlineShapes.Count == 0 && selection.ShapeRange.Count == 0)
            {
                toggleButton.Checked = false;
                return;
            }

            float height = 0f;
            if (selection.InlineShapes.Count > 0)
                height = selection.InlineShapes[1].Height;
            else if (selection.ShapeRange.Count > 0)
                height = selection.ShapeRange[1].Height;

            if (height <= 0)
            {
                toggleButton.Checked = false;
                return;
            }

            copiedHeight = height;
            lastHeightActivityTime = DateTime.Now;
            lastHeightSelectionHash = GetHeightSelectionHash(selection);
            StartHeightBrushMonitoring(app);
            app.StatusBar = $"高度刷已激活 - 高度: {height}磅 - 按ESC或点击空白区域退出";
        }

        public void DeactivateHeightBrush(Word.Application app)
        {
            heightBrushTimer?.Stop();
            heightBrushTimer?.Dispose();
            heightBrushTimer = null;
            lastHeightSelectionStart = null;
            lastHeightSelectionEnd = null;
            lastHeightSelectionHash = 0;
            try { app.StatusBar = ""; } catch { }
        }

        private void StartHeightBrushMonitoring(Word.Application app)
        {
            heightBrushTimer = new Timer();
            heightBrushTimer.Interval = 100;
            heightBrushTimer.Tick += (s, timerE) =>
            {
                // 注意：这里需要传入高度刷按钮的引用
                try
                {
                    var currentSelection = app.Selection;
                    int currentHash = GetHeightSelectionHash(currentSelection);

                    // 检查退出条件
                    if (currentHash != lastHeightSelectionHash)
                    {
                        bool isEmptyNow = currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                            (currentSelection.Type == Word.WdSelectionType.wdSelectionIP || currentSelection.Type == Word.WdSelectionType.wdSelectionNormal) &&
                            string.IsNullOrWhiteSpace(currentSelection.Text?.Replace("\r", ""));

                        TimeSpan timeSinceLastActivity = DateTime.Now - lastHeightActivityTime;
                        if (isEmptyNow && timeSinceLastActivity.TotalMilliseconds < 500) { DeactivateHeightBrush(app); return; }

                        if (currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                            currentSelection.Paragraphs.Count > 0 &&
                            string.IsNullOrWhiteSpace(currentSelection.Paragraphs[1].Range.Text.Replace("\r", "")))
                        { DeactivateHeightBrush(app); return; }
                    }

                    // 长时间停留在空白区域退出
                    if (currentSelection.InlineShapes.Count == 0 && currentSelection.ShapeRange.Count == 0 &&
                        string.IsNullOrWhiteSpace(currentSelection.Text?.Replace("\r", "")) &&
                        (DateTime.Now - lastHeightActivityTime).TotalMilliseconds > 1000)
                    { DeactivateHeightBrush(app); return; }

                    // 应用高度
                    bool hasImageSelection = currentSelection.InlineShapes.Count > 0 || currentSelection.ShapeRange.Count > 0;
                    if (hasImageSelection)
                    {
                        bool isNewSelection = false;
                        try
                        {
                            if (lastHeightSelectionStart == null || lastHeightSelectionEnd == null ||
                                !lastHeightSelectionStart.Equals(currentSelection.Start) || !lastHeightSelectionEnd.Equals(currentSelection.End))
                            {
                                isNewSelection = true;
                                lastHeightSelectionStart = currentSelection.Start;
                                lastHeightSelectionEnd = currentSelection.End;
                                lastHeightActivityTime = DateTime.Now;
                                lastHeightSelectionHash = currentHash;
                            }
                        }
                        catch { isNewSelection = true; lastHeightActivityTime = DateTime.Now; lastHeightSelectionHash = currentHash; }

                        if (isNewSelection && ApplyHeightToSelection(currentSelection))
                            app.StatusBar = $"高度刷: 已应用高度 {copiedHeight}磅 - 按ESC或点击空白区域退出";
                    }
                    else if (currentHash != lastHeightSelectionHash)
                    {
                        lastHeightActivityTime = DateTime.Now;
                        lastHeightSelectionHash = currentHash;
                    }
                }
                catch { }
            };
            heightBrushTimer.Start();
        }

        private int GetHeightSelectionHash(Word.Selection selection)
        {
            try
            {
                int hash = 0;
                hash ^= selection.Start.GetHashCode();
                hash ^= selection.End.GetHashCode();
                hash ^= selection.Type.GetHashCode();
                hash ^= selection.InlineShapes.Count.GetHashCode();
                hash ^= selection.ShapeRange.Count.GetHashCode();
                if (!string.IsNullOrEmpty(selection.Text))
                    hash ^= selection.Text.GetHashCode();
                return hash;
            }
            catch { return DateTime.Now.Millisecond; }
        }

        private bool ApplyHeightToSelection(Word.Selection selection)
        {
            bool applied = false;
            try
            {
                if (selection.InlineShapes.Count > 0)
                {
                    foreach (Word.InlineShape shape in selection.InlineShapes)
                    {
                        if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeChart ||
                            shape.Type == Word.WdInlineShapeType.wdInlineShapeSmartArt)
                        {
                            shape.Height = copiedHeight;
                            applied = true;
                        }
                    }
                }

                if (selection.ShapeRange.Count > 0)
                {
                    foreach (Word.Shape shape in selection.ShapeRange)
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoChart ||
                            shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                        {
                            shape.Height = copiedHeight;
                            applied = true;
                        }
                    }
                }
            }
            catch { }
            return applied;
        }
        #endregion

        #region 位图化相关方法
        public void ConvertToBitmap_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selection = app.Selection;

                // 检查是否选中了单个图形并转换
                if (selection.InlineShapes.Count == 1)
                    ConvertToBitmap(selection.InlineShapes[1], app);
                else if (selection.ShapeRange.Count == 1)
                    ConvertToBitmap(selection.ShapeRange[1], app);
            }
            catch
            {
                // 静默处理错误，不显示提示
            }
        }

        // 转换方法
        private void ConvertToBitmap(object shape, Word.Application app)
        {
            // 如果不是位图则转换
            if (!IsBitmap(shape))
            {
                // 统一处理：选择、复制、删除、粘贴
                if (shape is Word.InlineShape inlineShape)
                {
                    inlineShape.Select();
                    app.Selection.Copy();
                    inlineShape.Delete();
                }
                else if (shape is Word.Shape wordShape)
                {
                    wordShape.Select();
                    app.Selection.Copy();
                    wordShape.Delete();
                }
                app.Selection.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteBitmap);
            }
        }

        // 判断是否为位图
        private bool IsBitmap(object shape)
        {
            if (shape is Word.InlineShape inlineShape)
            {
                var type = inlineShape.Type;
                return (type == Word.WdInlineShapeType.wdInlineShapePicture ||
                        type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture) &&
                       inlineShape.PictureFormat != null;
            }
            else if (shape is Word.Shape wordShape)
            {
                var type = wordShape.Type;
                return (type == Microsoft.Office.Core.MsoShapeType.msoPicture ||
                        type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture) &&
                       wordShape.PictureFormat != null;
            }
            return false;
        }
        #endregion

        #region 清理方法
        public void Cleanup()
        {
            widthBrushTimer?.Stop();
            widthBrushTimer?.Dispose();
            widthBrushTimer = null;

            heightBrushTimer?.Stop();
            heightBrushTimer?.Dispose();
            heightBrushTimer = null;
        }
        #endregion
    }
}