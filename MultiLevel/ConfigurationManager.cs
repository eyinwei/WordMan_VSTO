using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordMan_VSTO
{
    /// <summary>
    /// 配置管理器 - 专门处理配置文件的加载、保存和数据处理
    /// </summary>
    public static class ConfigurationManager
    {
        #region 配置文件操作

        /// <summary>
        /// 保存配置到TXT文件（使用分号分割）
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="levelDataList">级别数据列表</param>
        /// <param name="currentLevels">当前级别数</param>
        public static void SaveConfigurationToFile(string filePath, List<LevelData> levelDataList, int currentLevels)
        {
            var content = new StringBuilder();
            // 第一行：列表级数信息
            content.AppendLine($"列表级数;{currentLevels}");
            // 第二行：列标题
            content.AppendLine("级别;编号样式;编号格式;编号缩进;文本缩进;编号之后类型;制表位位置;链接样式");
            
            // 只导出当前设置的级数（1到currentLevels）
            for (int i = 1; i <= currentLevels; i++)
            {
                var levelData = levelDataList.FirstOrDefault(x => x.Level == i);
                if (levelData != null)
                {
                    content.AppendLine($"{levelData.Level};{levelData.NumberStyle};{levelData.NumberFormat};" +
                                     $"{levelData.NumberIndent};{levelData.TextIndent};{levelData.AfterNumberType};" +
                                     $"{levelData.TabPosition};{levelData.LinkedStyle}");
                }
            }
            
            File.WriteAllText(filePath, content.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// 从TXT文件加载配置（使用分号分割）
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="levelDataList">输出的级别数据列表</param>
        /// <param name="currentLevels">输出的当前级别数</param>
        public static void LoadConfigurationFromFile(string filePath, out List<LevelData> levelDataList, out int currentLevels)
        {
            levelDataList = new List<LevelData>();
            currentLevels = 0; // 默认不显示任何级别
            
            if (!File.Exists(filePath))
                return;
            
            var lines = File.ReadAllLines(filePath, Encoding.UTF8);
            if (lines.Length <= 2) // 只有列表级数行和标题行
                return;
            
            // 第一行：读取列表级数信息
            if (lines.Length > 0)
            {
                var levelInfoParts = lines[0].Split(';');
                if (levelInfoParts.Length >= 2 && levelInfoParts[0] == "列表级数")
                {
                    if (int.TryParse(levelInfoParts[1], out int parsedLevels))
                    {
                        currentLevels = parsedLevels;
                    }
                }
            }
            
            // 从第三行开始读取数据（跳过列表级数行和标题行）
            for (int i = 2; i < lines.Length; i++)
            {
                var parts = lines[i].Split(';');
                if (parts.Length >= 8)
                {
                    var levelData = new LevelData
                    {
                        Level = int.Parse(parts[0]),
                        NumberStyle = parts[1],
                        NumberFormat = parts[2],
                        NumberIndent = decimal.Parse(parts[3]),
                        TextIndent = decimal.Parse(parts[4]),
                        AfterNumberType = parts[5],
                        TabPosition = decimal.Parse(parts[6]),
                        LinkedStyle = parts[7]
                    };
                    levelDataList.Add(levelData);
                }
            }
            
            // 如果没有从文件中读取到列表级数，则根据数据计算
            if (currentLevels == 0 && levelDataList.Count > 0)
            {
                currentLevels = levelDataList.Max(x => x.Level);
            }
        }


        #endregion

        #region 对话框操作

        /// <summary>
        /// 显示导入对话框
        /// </summary>
        /// <returns>选择的文件路径，如果取消则返回空字符串</returns>
        public static string ShowImportDialog()
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "TXT文件|*.txt|所有文件|*.*";
                openFileDialog.Title = "导入多级列表配置";
                openFileDialog.DefaultExt = "txt";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// 显示导出对话框
        /// </summary>
        /// <returns>选择的文件路径，如果取消则返回空字符串</returns>
        public static string ShowExportDialog()
        {
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "TXT文件|*.txt|所有文件|*.*";
                saveFileDialog.Title = "导出多级列表配置";
                saveFileDialog.DefaultExt = "txt";
                saveFileDialog.FileName = $"多级列表配置_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
            }
            return string.Empty;
        }

        #endregion


    }
}
