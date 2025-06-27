using Desktop.Models;
using Desktop.ViewModels;
using System.IO;
using System.Xml.Serialization;

namespace Desktop.Helpers
{
    public class XmlHelper
    {
        public static MixInfo LoadFromXml(string filePath)
        {
            if (!File.Exists(filePath)) return new MixInfo();

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                var serializer = new XmlSerializer(typeof(MixInfo));
                return (MixInfo)serializer.Deserialize(stream)!;
            }
        }

        public static void SaveToXml(string filePath, MixInfo data)
        {
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                var serializer = new XmlSerializer(typeof(MixInfo));
                serializer.Serialize(stream, data);
            }
        }
    }
}
