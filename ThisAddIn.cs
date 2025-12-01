using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordMan
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 监听选择变化，自动更新重复标题行按钮状态
            Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            try
            {
                // 无论是否在表格中，都更新重复标题行按钮状态
                // 当光标移出表格时，按钮状态会自动更新为未选中
                var ribbon = Globals.Ribbons.GetRibbon<MainRibbon>();
                if (ribbon != null)
                {
                    ribbon.UpdateRepeatHeaderRowsButtonState();
                }
            }
            catch
            {
                // 忽略错误，避免影响正常使用
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 清理资源
            try
            {
                // 移除事件监听
                if (Application != null)
                {
                    Application.WindowSelectionChange -= Application_WindowSelectionChange;
                }

                var ribbon = Globals.Ribbons.GetRibbon<MainRibbon>();
                if (ribbon != null)
                {
                    ribbon.Cleanup();
                }
            }
            catch
            {
                // 忽略清理时的错误，避免影响正常关闭
            }
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
