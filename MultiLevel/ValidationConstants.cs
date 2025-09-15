using System;

namespace WordMan_VSTO
{
    /// <summary>
    /// 验证常量 - 统一管理所有验证相关的常量
    /// </summary>
    public static class ValidationConstants
    {
        /// <summary>
        /// 有效的编号样式列表
        /// </summary>
        public static readonly string[] ValidNumberStyles = 
        {
            "1,2,3...", "01,02,03...", "A,B,C...", "a,b,c...", 
            "I,II,III...", "i,ii,iii...", "一,二,三...", "壹,贰,叁...", 
            "甲,乙,丙...", "正规编号"
        };

        /// <summary>
        /// 有效的编号之后类型列表
        /// </summary>
        public static readonly string[] ValidAfterNumberTypes = 
        {
            "无", "空格", "制表位"
        };

        /// <summary>
        /// 有效的链接样式列表
        /// </summary>
        public static readonly string[] ValidLinkedStyles = 
        {
            "无", "标题 1", "标题 2", "标题 3", "标题 4", 
            "标题 5", "标题 6", "标题 7", "标题 8", "标题 9"
        };

        /// <summary>
        /// 默认编号样式
        /// </summary>
        public const string DefaultNumberStyle = "1,2,3...";

        /// <summary>
        /// 默认编号之后类型
        /// </summary>
        public const string DefaultAfterNumberType = "空格";

        /// <summary>
        /// 默认链接样式
        /// </summary>
        public const string DefaultLinkedStyle = "无";

        /// <summary>
        /// 最小级别数
        /// </summary>
        public const int MinLevelCount = 1;

        /// <summary>
        /// 最大级别数
        /// </summary>
        public const int MaxLevelCount = 9;
    }
}
