using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace WordMan_VSTO.StylePane
{
    public static class StyleSerializationHelper
    {
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
            using (StreamWriter textWriter = new StreamWriter(filePath))
            {
                xmlSerializer.Serialize(textWriter, obj);
            }
        }

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
            using (StreamWriter textWriter = new StreamWriter(filePath))
            {
                xmlSerializer.Serialize(textWriter, list);
            }
        }

        public static T DeserializeFromXml<T>(string filePath) where T : class
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到指定的文件", filePath);
            }
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(T));
            using (StreamReader textReader = new StreamReader(filePath))
            {
                return xmlSerializer.Deserialize(textReader) as T;
            }
        }

        public static List<T> DeserializeListFromXml<T>(string filePath) where T : class
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("找不到指定的文件", filePath);
            }
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(List<T>));
            using (StreamReader textReader = new StreamReader(filePath))
            {
                return xmlSerializer.Deserialize(textReader) as List<T>;
            }
        }
    }
}
