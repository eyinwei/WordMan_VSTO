using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordMan_VSTO
{
    /// <summary>
    /// 配置管理器 - 处理多级列表配置的导入导出
    /// </summary>
    public static class ConfigurationManager
    {
        /// <summary>
        /// 保存配置到文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="levelDataList">级别数据列表</param>
        /// <param name="currentLevels">当前级别数</param>
        public static void SaveConfigurationToFile(string filePath, List<LevelData> levelDataList, int currentLevels)
        {
            try
            {
                var configData = new StringBuilder();
                
                // 添加CSV格式的文件头
                configData.AppendLine($"级数,{currentLevels},,,,,,");
                configData.AppendLine("级别,编号样式,编号格式,编号缩进,文本缩进,编号之后,制表位位置,链接样式");
                
                // 添加各级别数据
                for (int level = 1; level <= currentLevels; level++)
                {
                    var levelData = levelDataList[level - 1];
                    // 对包含逗号的字段进行转义
                    string numberStyle = EscapeCsvField(levelData.NumberStyle);
                    string numberFormat = EscapeCsvField(levelData.NumberFormat);
                    string afterNumberType = EscapeCsvField(levelData.AfterNumberType);
                    string linkedStyle = EscapeCsvField(levelData.LinkedStyle);
                    
                    configData.AppendLine($"{level},{numberStyle},{numberFormat},{levelData.NumberIndent},{levelData.TextIndent},{afterNumberType},{levelData.TabPosition},{linkedStyle}");
                }
                
                // 保存到文件
                File.WriteAllText(filePath, configData.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                throw new Exception($"保存配置文件失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 从文件加载配置
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="levelDataList">级别数据列表（输出）</param>
        /// <param name="currentLevels">当前级别数（输出）</param>
        public static void LoadConfigurationFromFile(string filePath, out List<LevelData> levelDataList, out int currentLevels)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException("配置文件不存在");
                }

                var lines = File.ReadAllLines(filePath, Encoding.UTF8);
                if (lines.Length == 0)
                {
                    throw new Exception("配置文件为空");
                }
                
                // 初始化输出参数
                levelDataList = new List<LevelData>();
                currentLevels = 4; // 默认值
                
                // 初始化级别数据
                for (int i = 1; i <= 9; i++)
                {
                    levelDataList.Add(new LevelData
                    {
                        Level = i,
                        NumberStyle = "1,2,3...",
                        NumberFormat = "",
                        NumberIndent = 0m,
                        TextIndent = 0m,
                        AfterNumberType = "空格",
                        TabPosition = 0m,
                        LinkedStyle = "无"
                    });
                }
                
                foreach (var line in lines)
                {
                    // 跳过空行
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    
                    // 解析CSV行
                    var parts = ParseCsvLine(line);
                    if (parts.Length < 2)
                        continue;
                    
                    // 处理级别数
                    if (parts[0] == "级数")
                    {
                        if (int.TryParse(parts[1], out int levelCount))
                        {
                            currentLevels = ValidationHelper.ValidateLevelCount(levelCount);
                        }
                        else
                        {
                            throw new Exception($"无效的级别数：{parts[1]}，必须是1-9之间的整数");
                        }
                        continue;
                    }
                    
                    // 跳过标题行
                    if (parts[0] == "级别")
                        continue;
                    
                    // 解析级别数据（格式：级别,编号样式,编号格式,编号缩进,文本缩进,编号之后,制表位位置,链接样式）
                    if (int.TryParse(parts[0], out int level) && level >= 1 && level <= 9)
                    {
                        if (parts.Length >= 8)
                        {
                            var levelData = levelDataList[level - 1];
                            
                            // 设置文本字段，验证有效性
                            levelData.NumberStyle = ValidationHelper.ValidateNumberStyle(parts[1] ?? "");
                            levelData.NumberFormat = parts[2] ?? "";
                            levelData.AfterNumberType = ValidationHelper.ValidateAfterNumberType(parts[5] ?? "");
                            levelData.LinkedStyle = ValidationHelper.ValidateLinkedStyle(parts[7] ?? "");
                            
                            // 设置数值字段，提供默认值
                            if (decimal.TryParse(parts[3], out decimal numberIndent) && numberIndent >= 0)
                                levelData.NumberIndent = numberIndent;
                            else
                                levelData.NumberIndent = 0;
                                
                            if (decimal.TryParse(parts[4], out decimal textIndent) && textIndent >= 0)
                                levelData.TextIndent = textIndent;
                            else
                                levelData.TextIndent = 0;
                                
                            if (decimal.TryParse(parts[6], out decimal tabPosition) && tabPosition >= 0)
                                levelData.TabPosition = tabPosition;
                            else
                                levelData.TabPosition = 0;
                            
                            levelDataList[level - 1] = levelData;
                        }
                        else
                        {
                            throw new Exception($"级别 {level} 的数据不完整，需要8个字段，实际只有 {parts.Length} 个");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"加载配置文件失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 显示导入对话框
        /// </summary>
        /// <returns>选择的文件路径，如果取消则返回null</returns>
        public static string ShowImportDialog()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "CSV配置文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                openFileDialog.Title = "导入多级列表配置";
                openFileDialog.DefaultExt = "csv";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return null;
        }

        /// <summary>
        /// 显示导出对话框
        /// </summary>
        /// <returns>选择的文件路径，如果取消则返回null</returns>
        public static string ShowExportDialog()
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "CSV配置文件 (*.csv)|*.csv|所有文件 (*.*)|*.*";
                saveFileDialog.Title = "导出多级列表配置";
                saveFileDialog.DefaultExt = "csv";
                saveFileDialog.FileName = "多级列表配置.csv";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
            }
            return null;
        }

        /// <summary>
        /// 转义CSV字段，处理包含逗号、引号或换行符的字段
        /// </summary>
        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "";
            
            // 如果字段包含逗号、引号或换行符，需要用引号包围并转义内部引号
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }
            
            return field;
        }

        /// <summary>
        /// 解析CSV行，正确处理引号包围的字段
        /// </summary>
        private static string[] ParseCsvLine(string line)
        {
            var fields = new List<string>();
            var currentField = new StringBuilder();
            bool inQuotes = false;
            
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // 转义的引号
                        currentField.Append('"');
                        i++; // 跳过下一个引号
                    }
                    else
                    {
                        // 开始或结束引号
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    // 字段分隔符
                    fields.Add(currentField.ToString());
                    currentField.Clear();
                }
                else
                {
                    currentField.Append(c);
                }
            }
            
            // 添加最后一个字段
            fields.Add(currentField.ToString());
            
            return fields.ToArray();
        }

    }
}
