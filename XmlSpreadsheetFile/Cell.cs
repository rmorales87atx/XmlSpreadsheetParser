using System;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("Cell", Namespace = XmlNamespaces.Spreadsheet)]
    public class Cell {
        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string StyleID { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string Index { get; set; }

        [XmlElement]
        public Data Data { get; set; }

        public override string ToString() => Data.Value;
    }

}
