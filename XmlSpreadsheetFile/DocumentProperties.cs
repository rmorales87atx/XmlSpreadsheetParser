using System;
using System.Xml;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("DocumentProperties")]
    public class DocumentProperties {
        [XmlElement]
        public string Author { get; set; }

        public string LastAuthor { get; set; }

        public DateTime LastPrinted { get; set; }

        public DateTime Created { get; set; }

        public DateTime LastSaved { get; set; }

        public string Version { get; set; }
    }

}
