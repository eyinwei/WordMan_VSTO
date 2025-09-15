using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace WordMan_VSTO
{
    /// <summary>
    /// 样式序列化辅助类
    /// 用于保存和加载样式设置到XML文件
    /// </summary>
    public static class StyleSerializationHelper
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
    }
}
