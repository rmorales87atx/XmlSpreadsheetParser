using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {

    [Serializable, XmlRoot("Data", Namespace = XmlNamespaces.Spreadsheet)]
    public class Data {
        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string Type { get; set; }

        [XmlText]
        public string Value { get; set; }

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns = new XmlSerializerNamespaces();

        public Data()
        {
            xmlns.Add("ss", XmlNamespaces.Spreadsheet);
        }

        public override string ToString() => Value;
    }

}
