using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("Worksheet", Namespace = XmlNamespaces.Spreadsheet)]
    [XmlType(Namespace = XmlNamespaces.Spreadsheet)]
    public class Worksheet {
        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string Name { get; set; }

        [XmlElement("Table", Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public Table Table { get; set; }

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns = new XmlSerializerNamespaces();

        public Worksheet()
        {
            xmlns.Add("ss", XmlNamespaces.Spreadsheet);
        }

        public override string ToString() => Name;
    }

}
