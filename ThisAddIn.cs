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

        #region 全局工具方法
        /// <summary>
        /// 执行操作并将其封装为一个撤销步骤
        /// </summary>
        /// <param name="undoRecordName">撤销记录的名称，将显示在撤销历史中</param>
        /// <param name="action">要执行的操作</param>
        public void ExecuteWithUndoRecord(string undoRecordName, System.Action action)
        {
            if (action == null) return;

            Word.UndoRecord undoRecord = null;
            try
            {
                var doc = Application.ActiveDocument;
                
                if (doc == null) return;
                
                // 开始自定义撤销记录
                undoRecord = doc.Application.UndoRecord;
                undoRecord.StartCustomRecord(undoRecordName);
                
                // 执行操作
                action();
            }
            finally
            {
                // 确保撤销记录已结束（无论是否出错）
                if (undoRecord != null)
                {
                    try
                    {
                        undoRecord.EndCustomRecord();
                    }
                    catch { }
                }
            }
        }
        #endregion

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
