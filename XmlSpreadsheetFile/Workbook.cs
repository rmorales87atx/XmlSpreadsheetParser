using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("Workbook", Namespace = XmlNamespaces.Spreadsheet)]
    public class Workbook {
        [XmlElement("DocumentProperties", Namespace = XmlNamespaces.Office, Form = XmlSchemaForm.None)]
        public DocumentProperties DocumentProperties { get; set; }

        [XmlElement("ExcelWorkbook", Namespace = XmlNamespaces.Excel, Form = XmlSchemaForm.None)]
        public WorkbookProperties WorkbookProperties { get; set; }

        [XmlArray("Styles", Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        [XmlArrayItem("Style", Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public List<Style> Styles { get; set; }

        [XmlElement("Worksheet", Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.None)]
        public List<Worksheet> Worksheets { get; set; }

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns = new XmlSerializerNamespaces();

        public Workbook()
        {
            xmlns.Add("o", XmlNamespaces.Office);
            xmlns.Add("x", XmlNamespaces.Excel);
            xmlns.Add("ss", XmlNamespaces.Spreadsheet);
        }

        public static Workbook LoadFromFile(string path)
        {
            using (var reader = new StreamReader(path)) {
                var serializer = new XmlSerializer(typeof(Workbook));
                return (Workbook)serializer.Deserialize(reader);
            }
        }

        public void SaveToFile(string path)
        {
            DocumentProperties.LastSaved = DateTime.UtcNow;
            using (var writer = new StreamWriter(path)) {
                var serializer = new XmlSerializer(typeof(Workbook));
                serializer.Serialize(writer, this);
            }
        }
    }

}
