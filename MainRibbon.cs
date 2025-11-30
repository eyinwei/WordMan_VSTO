using Microsoft.Office.Tools.Ribbon;
using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using WordMan.SplitAndMerge;
using WordMan.MultiLevel;
using WordMan;
using static WordMan.CaptionManager;

namespace WordMan
{
    public partial class MainRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private ImageProcessor imageProcessor = new ImageProcessor();
        private TextProcessor textProcessor = new TextProcessor();
        private TableProcessor tableProcessor = new TableProcessor();
        private CaptionManager captionManager = new CaptionManager();
        private DocumentProcessor documentProcessor = new DocumentProcessor();

        #region 文本处理组
        // Word 内置功能
        private void 清除格式_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.ClearFormatting();
        }

        private void 格式刷_Click(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            textProcessor.FormatPainter_Click(toggleButton);
        }

        private void 只留文本_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.PasteTextOnly();
        }

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

        private void 仿宋替换_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.ReplaceFangSongGB2312ToFangSong();
        }

        private void 楷体替换_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.ReplaceKaiTiGB2312ToKaiTi();
        }

        private void 方正小标宋替换_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.ReplaceFZXBSToHeiTi();
        }

        private void 数字替换_Click(object sender, RibbonControlEventArgs e)
        {
            textProcessor.ReplaceAllToTimesNewRoman();
        }
        #endregion

        #region 表格处理组
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

        private void 重复标题行_Click(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            tableProcessor.RepeatHeaderRows(toggleButton);
        }
        #endregion

        #region 题注与引用组
        private void 图注样式1_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetPictureStyle(图注样式1, 图注样式2, 图注样式3, CaptionNumberStyle.Arabic);
        }

        private void 图注样式2_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetPictureStyle(图注样式2, 图注样式1, 图注样式3, CaptionNumberStyle.Dash);
        }

        private void 图注样式3_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetPictureStyle(图注样式3, 图注样式1, 图注样式2, CaptionNumberStyle.Dot);
        }

        private void 图编号_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.InsertPictureNumber();
        }

        private void 表注样式1_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetTableStyle(表注样式1, 表注样式2, 表注样式3, CaptionNumberStyle.Arabic);
        }

        private void 表注样式2_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetTableStyle(表注样式2, 表注样式1, 表注样式3, CaptionNumberStyle.Dash);
        }

        private void 表注样式3_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetTableStyle(表注样式3, 表注样式1, 表注样式2, CaptionNumberStyle.Dot);
        }

        private void 表编号_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.InsertTableNumber();
        }

        private void 公式样式1_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetFormulaStyle(公式样式1, 公式样式2, 公式样式3, FormulaNumberStyle.Parenthesis1);
        }

        private void 公式样式2_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetFormulaStyle(公式样式2, 公式样式1, 公式样式3, FormulaNumberStyle.Parenthesis1_1);
        }

        private void 公式样式3_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.SetFormulaStyle(公式样式3, 公式样式1, 公式样式2, FormulaNumberStyle.Parenthesis1_1dot);
        }

        private void 式编号_Click(object sender, RibbonControlEventArgs e)
        {
            captionManager.InsertFormulaNumber();
        }

        private void 交叉引用_Click(object sender, RibbonControlEventArgs e)
        {
            var toggleButton = sender as Microsoft.Office.Tools.Ribbon.RibbonToggleButton;
            captionManager.ToggleCrossReferenceMode(toggleButton);
        }
        #endregion

        #region 图片处理组
        private void 宽度刷_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.WidthBrush_Click(sender, e, 宽度刷);
        }

        private void 高度刷_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.HeightBrush_Click(sender, e, 高度刷);
        }

        private void 位图化_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.ConvertToBitmap_Click(sender, e);
        }

        private void 导出图片_Click(object sender, RibbonControlEventArgs e)
        {
            imageProcessor.ExportImage_Click(sender, e);
        }

        public void Cleanup()
        {
            imageProcessor.Cleanup();
        }
        #endregion

        #region 全文处理组
        private void TypesettingButton_Click(object sender, RibbonControlEventArgs e)
        {
            TypesettingTaskPane.TriggerShowOrHide();
        }

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

        private void 域名高亮_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.HighlightFields(true);
        }

        private void 取消高亮_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.HighlightFields(false);
        }

        private void 上标_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.SetFieldSuperscript(true);
        }

        private void 正常_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.SetFieldSuperscript(false);
        }

        private void 另存PDF_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.ExportToPDF();
        }

        private void 版本_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.ShowVersion();
        }

        private void 文档合并_Click(object sender, RibbonControlEventArgs e)
        {
            var merger = new DocumentMerger((Word.Application)Globals.ThisAddIn.Application);
            merger.ShowMergeDialog();
        }

        private void 文档拆分_Click(object sender, RibbonControlEventArgs e)
        {
            var splitter = new DocumentSplitter(Globals.ThisAddIn.Application);
            splitter.ShowSplitDialog();
        }

        private void 公开_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.AddSecurityLevel("公开");
        }

        private void 内部_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.AddSecurityLevel("内部★");
        }

        private void 移除密级_Click(object sender, RibbonControlEventArgs e)
        {
            documentProcessor.RemoveSecurityLevelFromCurrentPage();
        }
        #endregion
    }
}
