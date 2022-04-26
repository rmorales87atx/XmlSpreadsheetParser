using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("Column", Namespace = XmlNamespaces.Spreadsheet)]
    public class Column {
        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string Index { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string AutoFitWidth { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string Width { get; set; }
    }

}
