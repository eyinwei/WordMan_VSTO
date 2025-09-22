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
        /// 保存配置到CSV文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="levelDataList">级别数据列表</param>
        /// <param name="currentLevels">当前级别数</param>
        public static void SaveConfigurationToFile(string filePath, List<LevelData> levelDataList, int currentLevels)
        {
            var csv = new StringBuilder();
            csv.AppendLine("Level,NumberStyle,NumberFormat,NumberIndent,TextIndent,AfterNumberType,TabPosition,LinkedStyle");
            
            foreach (var levelData in levelDataList)
            {
                csv.AppendLine($"{levelData.Level},{levelData.NumberStyle},{levelData.NumberFormat}," +
                             $"{levelData.NumberIndent},{levelData.TextIndent},{levelData.AfterNumberType}," +
                             $"{levelData.TabPosition},{levelData.LinkedStyle}");
            }
            
            File.WriteAllText(filePath, csv.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// 从CSV文件加载配置
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
            if (lines.Length <= 1) // 只有标题行
                return;
            
            for (int i = 1; i < lines.Length; i++) // 跳过标题行
            {
                var parts = lines[i].Split(',');
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
            
            currentLevels = levelDataList.Count > 0 ? levelDataList.Max(x => x.Level) : 0;
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
                openFileDialog.Filter = "CSV文件|*.csv|所有文件|*.*";
                openFileDialog.Title = "导入多级列表配置";
                openFileDialog.DefaultExt = "csv";
                
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
                saveFileDialog.Filter = "CSV文件|*.csv|文本文件|*.txt|所有文件|*.*";
                saveFileDialog.Title = "导出多级列表配置";
                saveFileDialog.DefaultExt = "csv";
                saveFileDialog.FileName = $"多级列表配置_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                
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
