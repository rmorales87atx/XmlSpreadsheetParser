using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("Style", Namespace = XmlNamespaces.Spreadsheet)]
    public class Style {
        [XmlElement(Namespace = XmlNamespaces.Spreadsheet)]
        public object Alignment { get; set; }

        [XmlElement(Namespace = XmlNamespaces.Spreadsheet)]
        public object Borders { get; set; }

        [XmlElement(Namespace = XmlNamespaces.Spreadsheet)]
        public object Font { get; set; }

        [XmlElement(Namespace = XmlNamespaces.Spreadsheet)]
        public object Interior { get; set; }

        [XmlElement(Namespace = XmlNamespaces.Spreadsheet)]
        public object NumberFormat { get; set; }

        [XmlElement(Namespace = XmlNamespaces.Spreadsheet)]
        public object Protection { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string ID { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string Name { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string Parent { get; set; }

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns = new XmlSerializerNamespaces();

        public Style()
        {
            xmlns.Add("ss", XmlNamespaces.Spreadsheet);
        }

        public override string ToString() => ID;
    }
}
