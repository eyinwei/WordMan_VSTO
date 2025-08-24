using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    public partial class MainRibbon
    {

        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void 去除断行_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection;
            if (sel == null || sel.Range == null || string.IsNullOrEmpty(sel.Range.Text))
                return;

            Word.Range rng = sel.Range.Duplicate;
            string text = rng.Text;

            // 判断末尾是否有回车
            bool endsWithReturn = text.EndsWith("\r");

            // 如果结尾有回车，先排除最后一个回车后再处理
            int processLength = endsWithReturn ? text.Length - 1 : text.Length;
            Word.Range processRange = rng.Duplicate;
            processRange.End = processRange.Start + processLength;

            // 替换所有回车
            processRange.Find.ClearFormatting();
            processRange.Find.Replacement.ClearFormatting();
            processRange.Find.Text = "^013"; // 回车
            processRange.Find.Replacement.Text = "";
            processRange.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);

            // 替换所有软回车
            processRange.Find.Text = "^11"; // 手动换行(软回车)
            processRange.Find.Replacement.Text = "";
            processRange.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);

            // 这样可一键撤销，且格式不会丢失
        }
        private void 去除空格_Click(object sender, RibbonControlEventArgs e)
        {
            var rng = Globals.ThisAddIn.Application.ActiveDocument.Content;
            rng.Find.Execute(" ", ReplaceWith: "", Replace: Word.WdReplace.wdReplaceAll);
            rng.Find.Execute("　", ReplaceWith: "", Replace: Word.WdReplace.wdReplaceAll);
        }

        private void 去除空行_Click(object sender, RibbonControlEventArgs e)
        {
            // 获取Word应用程序对象
            var app = Globals.ThisAddIn.Application;

            // 获取当前选区
            Word.Range rng = app.Selection.Range;

            // 从后往前遍历选区内的所有段落
            for (int i = rng.Paragraphs.Count; i >= 1; i--)
            {
                Word.Paragraph para = rng.Paragraphs[i];
                // 去除回车、换行、空格、全角空格、Tab等
                string text = para.Range.Text.Trim('\r', '\n', ' ', '\t', '　');
                if (string.IsNullOrEmpty(text))
                {
                    para.Range.Delete();
                }
            }
        }

        private void 去除缩进_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.Selection;
            if (selection != null)
            {
                var paraFormat = selection.ParagraphFormat;

                // 先清除首行缩进（字符和磅）
                paraFormat.CharacterUnitFirstLineIndent = 0f;
                paraFormat.FirstLineIndent = 0f;

                // 再清除段落左缩进（字符和磅）
                paraFormat.CharacterUnitLeftIndent = 0f;
                paraFormat.LeftIndent = 0f;

                // 可选：右缩进也清零（防止有些文档右缩进影响排版）
                paraFormat.CharacterUnitRightIndent = 0f;
                paraFormat.RightIndent = 0f;
            }
        }

        private void 缩进2字符_Click(object sender, RibbonControlEventArgs e)
        {
            var selection = Globals.ThisAddIn.Application.Selection;
            if (selection != null)
            {
                var paraFormat = selection.ParagraphFormat;
                paraFormat.CharacterUnitFirstLineIndent = 2f;
            }
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
    }
}



