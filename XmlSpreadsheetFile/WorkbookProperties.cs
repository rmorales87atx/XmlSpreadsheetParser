using System;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("ExcelWorkbook")]
    public class WorkbookProperties {
        public string WindowHeight { get; set; }

        public string WindowWidth { get; set; }

        public string WindowTopX { get; set; }

        public string WindowTopY { get; set; }

        public XmlBoolean ProtectStructure { get; set; }

        public XmlBoolean ProtectWindows { get; set; }
    }

}
