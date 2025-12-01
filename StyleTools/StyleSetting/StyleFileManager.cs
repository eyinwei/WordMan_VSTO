using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordMan
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
            
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
                using (StreamWriter textWriter = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
                {
                    xmlSerializer.Serialize(textWriter, obj);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"序列化对象到XML文件失败：{ex.Message}", ex);
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
            
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<T>));
                using (StreamWriter textWriter = new StreamWriter(filePath, false, System.Text.Encoding.UTF8))
                {
                    xmlSerializer.Serialize(textWriter, list);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"序列化列表到XML文件失败：{ex.Message}", ex);
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
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException("文件路径不能为空", "filePath");
            }
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到指定的文件", filePath);
            }
            
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
                using (StreamReader textReader = new StreamReader(filePath, System.Text.Encoding.UTF8))
                {
                    var result = xmlSerializer.Deserialize(textReader) as T;
                    if (result == null)
                    {
                        throw new InvalidOperationException("反序列化结果为空");
                    }
                    return result;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"从XML文件反序列化对象失败：{ex.Message}", ex);
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
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException("文件路径不能为空", "filePath");
            }
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到指定的文件", filePath);
            }
            
            try
            {
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<T>));
                using (StreamReader textReader = new StreamReader(filePath, System.Text.Encoding.UTF8))
                {
                    var result = xmlSerializer.Deserialize(textReader) as List<T>;
                    return result ?? new List<T>();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"从XML文件反序列化列表失败：{ex.Message}", ex);
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

    }
}
