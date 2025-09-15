using System.Linq;

namespace WordMan_VSTO
{
    /// <summary>
    /// 验证辅助类 - 统一管理所有验证逻辑
    /// </summary>
    public static class ValidationHelper
    {
        /// <summary>
        /// 验证编号样式
        /// </summary>
        /// <param name="style">要验证的样式</param>
        /// <returns>有效的样式，如果无效则返回默认值</returns>
        public static string ValidateNumberStyle(string style)
        {
            if (string.IsNullOrEmpty(style))
                return ValidationConstants.DefaultNumberStyle;
                
            if (ValidationConstants.ValidNumberStyles.Contains(style))
                return style;
                
            return ValidationConstants.DefaultNumberStyle;
        }

        /// <summary>
        /// 验证编号之后类型
        /// </summary>
        /// <param name="type">要验证的类型</param>
        /// <returns>有效的类型，如果无效则返回默认值</returns>
        public static string ValidateAfterNumberType(string type)
        {
            if (string.IsNullOrEmpty(type))
                return ValidationConstants.DefaultAfterNumberType;
                
            if (ValidationConstants.ValidAfterNumberTypes.Contains(type))
                return type;
                
            return ValidationConstants.DefaultAfterNumberType;
        }

        /// <summary>
        /// 验证链接样式
        /// </summary>
        /// <param name="style">要验证的样式</param>
        /// <returns>有效的样式，如果无效则返回默认值</returns>
        public static string ValidateLinkedStyle(string style)
        {
            if (string.IsNullOrEmpty(style))
                return ValidationConstants.DefaultLinkedStyle;
                
            if (ValidationConstants.ValidLinkedStyles.Contains(style))
                return style;
                
            return ValidationConstants.DefaultLinkedStyle;
        }

        /// <summary>
        /// 验证级别数
        /// </summary>
        /// <param name="levelCount">要验证的级别数</param>
        /// <returns>有效的级别数，如果无效则返回默认值</returns>
        public static int ValidateLevelCount(int levelCount)
        {
            if (levelCount >= ValidationConstants.MinLevelCount && levelCount <= ValidationConstants.MaxLevelCount)
                return levelCount;
                
            return ValidationConstants.MinLevelCount;
        }
    }
}
