using System;

namespace WordMan_VSTO
{
    /// <summary>
    /// 级别数据结构
    /// </summary>
    public class LevelData
    {
        public int Level { get; set; }
        public string NumberStyle { get; set; }
        public string NumberFormat { get; set; }
        public decimal NumberIndent { get; set; }
        public decimal TextIndent { get; set; }
        public string AfterNumberType { get; set; } // 编号之后类型：无、空格、制表位
        public decimal TabPosition { get; set; } // 制表位位置
        public string LinkedStyle { get; set; }
    }

    /// <summary>
    /// 级别数据事件参数
    /// </summary>
    public class LevelDataEventArgs : EventArgs
    {
        public LevelData LevelData { get; set; }
        
        public LevelDataEventArgs(LevelData levelData)
        {
            LevelData = levelData;
        }
    }

    /// <summary>
    /// 输入框值结构体
    /// </summary>
    public struct InputValues
    {
        public decimal NumberIndent { get; set; }
        public decimal TextIndent { get; set; }
        public decimal TabPosition { get; set; }
    }
}
