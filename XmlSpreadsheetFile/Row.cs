using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("Row")]
    public class Row {
        [XmlElement("Cell")]
        public List<Cell> Cells { get; set; }

        public override string ToString() => String.Join("\t", from x in Cells select x.Data.ToString());
    }

}
