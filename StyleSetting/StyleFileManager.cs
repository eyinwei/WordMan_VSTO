using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式文件管理器
    /// 用于保存和加载样式设置到XML文件，以及处理文件导入导出
    /// </summary>
    public static class StyleFileManager
    {
        /// <summary>
        /// 序列化对象到XML文件
        /// </summary>
        /// <typeparam name="T">要序列化的对象类型</typeparam>
        /// <param name="obj">要序列化的对象</param>
        /// <param name="filePath">文件路径</param>
        public static void SerializeToXml<T>(T obj, string filePath) where T : class
        {
            if (obj == null)
            {
                throw new ArgumentNullException("obj");
            }
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException("文件路径不能为空", "filePath");
            }
            
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
            using (StreamWriter textWriter = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                xmlSerializer.Serialize(textWriter, obj);
            }
        }

        /// <summary>
        /// 序列化列表到XML文件
        /// </summary>
        /// <typeparam name="T">要序列化的对象类型</typeparam>
        /// <param name="list">要序列化的列表</param>
        /// <param name="filePath">文件路径</param>
        public static void SerializeListToXml<T>(List<T> list, string filePath) where T : class
        {
            if (list == null)
            {
                throw new ArgumentNullException("list");
            }
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException("文件路径不能为空", "filePath");
            }
            
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<T>));
            using (StreamWriter textWriter = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
            {
                xmlSerializer.Serialize(textWriter, list);
            }
        }

        /// <summary>
        /// 从XML文件反序列化对象
        /// </summary>
        /// <typeparam name="T">要反序列化的对象类型</typeparam>
        /// <param name="filePath">文件路径</param>
        /// <returns>反序列化后的对象</returns>
        public static T DeserializeFromXml<T>(string filePath) where T : class
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到指定的文件", filePath);
            }
            
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
            using (StreamReader textReader = new StreamReader(filePath, System.Text.Encoding.UTF8))
            {
                return xmlSerializer.Deserialize(textReader) as T;
            }
        }

        /// <summary>
        /// 从XML文件反序列化列表
        /// </summary>
        /// <typeparam name="T">要反序列化的对象类型</typeparam>
        /// <param name="filePath">文件路径</param>
        /// <returns>反序列化后的列表</returns>
        public static List<T> DeserializeListFromXml<T>(string filePath) where T : class
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到指定的文件", filePath);
            }
            
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<T>));
            using (StreamReader textReader = new StreamReader(filePath, System.Text.Encoding.UTF8))
            {
                return xmlSerializer.Deserialize(textReader) as List<T>;
            }
        }

        /// <summary>
        /// 显示保存文件对话框
        /// </summary>
        /// <param name="defaultFileName">默认文件名</param>
        /// <param name="filter">文件过滤器</param>
        /// <returns>选择的文件路径，如果取消则返回null</returns>
        public static string ShowSaveFileDialog(string defaultFileName = "样式设置", string filter = "XML文件|*.xml|所有文件|*.*")
        {
            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = filter;
                saveFileDialog.FileName = defaultFileName;
                saveFileDialog.DefaultExt = "xml";
                saveFileDialog.AddExtension = true;
                
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
            }
            return null;
        }

        /// <summary>
        /// 显示打开文件对话框
        /// </summary>
        /// <param name="filter">文件过滤器</param>
        /// <returns>选择的文件路径，如果取消则返回null</returns>
        public static string ShowOpenFileDialog(string filter = "XML文件|*.xml|所有文件|*.*")
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = filter;
                openFileDialog.DefaultExt = "xml";
                
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return null;
        }

        /// <summary>
        /// 保存样式设置到文件
        /// </summary>
        /// <param name="styleSettings">样式设置字典</param>
        /// <param name="filePath">文件路径</param>
        public static void SaveStyleSettings(Dictionary<string, Hashtable> styleSettings, string filePath)
        {
            try
            {
                // 直接序列化字典
                SerializeToXml(styleSettings, filePath);
            }
            catch (Exception ex)
            {
                throw new Exception($"保存样式设置失败：{ex.Message}", ex);
            }
        }

        /// <summary>
        /// 从文件加载样式设置
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>样式设置字典</returns>
        public static Dictionary<string, Hashtable> LoadStyleSettings(string filePath)
        {
            try
            {
                // 直接反序列化字典
                return DeserializeFromXml<Dictionary<string, Hashtable>>(filePath);
            }
            catch (Exception ex)
            {
                throw new Exception($"加载样式设置失败：{ex.Message}", ex);
            }
        }

        /// <summary>
        /// 导出当前文档样式到文件
        /// </summary>
        /// <param name="app">Word应用程序对象</param>
        /// <param name="filePath">文件路径</param>
        public static void ExportDocumentStyles(Microsoft.Office.Interop.Word.Application app, string filePath)
        {
            try
            {
                var doc = app.ActiveDocument;
                var styleSettings = new Dictionary<string, Hashtable>();
                
                // 获取所有样式
                var styleNames = new[] { "标题 1", "标题 2", "标题 3", "标题 4", "标题 5", "标题 6", "正文", "题注", "表内文字" };
                
                foreach (var styleName in styleNames)
                {
                    try
                    {
                        var style = doc.Styles[styleName];
                        var settings = new Hashtable();
                        
                        // 读取样式属性
                        settings["cnFont"] = style.Font.Name;
                        settings["enFont"] = style.Font.NameAscii;
                        settings["fontSize"] = style.Font.Size;
                        settings["isBold"] = style.Font.Bold == 1;
                        settings["alignment"] = style.ParagraphFormat.Alignment;
                        settings["spaceBefore"] = app.PointsToCentimeters(style.ParagraphFormat.SpaceBefore);
                        settings["spaceAfter"] = app.PointsToCentimeters(style.ParagraphFormat.SpaceAfter);
                        settings["lineSpacing"] = style.ParagraphFormat.LineSpacing;
                        
                        styleSettings[styleName] = settings;
                    }
                    catch
                    {
                        // 样式不存在时跳过
                        continue;
                    }
                }
                
                SaveStyleSettings(styleSettings, filePath);
            }
            catch (Exception ex)
            {
                throw new Exception($"导出文档样式失败：{ex.Message}", ex);
            }
        }
    }
}
