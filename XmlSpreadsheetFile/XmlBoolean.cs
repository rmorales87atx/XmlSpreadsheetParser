using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    public struct XmlBoolean : IXmlSerializable {
        private bool _value;

        /// <summary>
        /// Allow implicit cast to a real bool
        /// </summary>
        /// <param name="yn">Value to cast to bool</param>
        public static implicit operator bool(XmlBoolean yn)
        {
            return yn._value;
        }

        /// <summary>
        /// Allow implicit cast from a real bool
        /// </summary>
        /// <param name="b">Value to cash to y/n</param>
        public static implicit operator XmlBoolean(bool b)
        {
            return new XmlBoolean { _value = b };
        }

        /// <summary>
        /// Reads a value from XML
        /// </summary>
        /// <param name="reader">XML reader to read</param>
        public void ReadXml(XmlReader reader)
        {
            _value = reader.ReadElementContentAsString().ToLowerInvariant().Equals("true");
        }

        /// <summary>
        /// Writes the value to XML
        /// </summary>
        /// <param name="writer">XML writer to write to</param>
        public void WriteXml(XmlWriter writer)
        {
            writer.WriteString(this.ToString());
        }

        public override string ToString() => _value ? "true" : "false";

        XmlSchema IXmlSerializable.GetSchema()
        {
            return null;
        }
    }
}
