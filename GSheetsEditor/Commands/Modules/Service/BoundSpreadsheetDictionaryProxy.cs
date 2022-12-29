using Newtonsoft.Json.Linq;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace GSheetsEditor.Commands.Modules.Service
{
    [XmlRoot("dictionary")]
    internal class BoundSpreadsheetDictionaryProxy : Dictionary<long, SpreadsheetsCollection>, IXmlSerializable
    {
        public BoundSpreadsheetDictionaryProxy() : base()
        {

        }

        public XmlSchema? GetSchema()
        {
            return null;
        }

        public void ReadXml(XmlReader reader)
        {
            XmlSerializer keySerializer = new XmlSerializer(typeof(long));
            XmlSerializer valueSerializer = new XmlSerializer(typeof(SpreadsheetsCollection));

            bool wasEmpty = reader.IsEmptyElement;
            reader.Read();

            if (wasEmpty)
                return;

            while (reader.NodeType != XmlNodeType.EndElement)
            {
                reader.ReadStartElement("item");

                reader.ReadStartElement("key");
                var key = (long)keySerializer.Deserialize(reader);
                reader.ReadEndElement();

                reader.ReadStartElement("value");
                var value = (SpreadsheetsCollection)valueSerializer.Deserialize(reader);
                reader.ReadEndElement();

                Add(key, value);

                reader.ReadEndElement();
                reader.MoveToContent();
            }
            reader.ReadEndElement();
            reader.Dispose();
        }

        public void WriteXml(XmlWriter writer)
        {
            XmlSerializer keySerializer = new XmlSerializer(typeof(long));
            XmlSerializer valueSerializer = new XmlSerializer(typeof(SpreadsheetsCollection));

            foreach (var key in Keys)
            {
                writer.WriteStartElement("item");

                writer.WriteStartElement("key");
                keySerializer.Serialize(writer, key);
                writer.WriteEndElement();

                writer.WriteStartElement("value");
                var value = this[key];
                valueSerializer.Serialize(writer, value);
                writer.WriteEndElement();

                writer.WriteEndElement();
            }
            writer.Dispose();
        }
    }
}
