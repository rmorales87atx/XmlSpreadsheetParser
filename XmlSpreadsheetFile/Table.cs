using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Serialization;

namespace XmlSpreadsheetFile {
    [Serializable, XmlRoot("Table", Namespace = XmlNamespaces.Spreadsheet)]
    public class Table {
        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string ExpandedColumnCount { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string ExpandedRowCount { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Excel, Form = XmlSchemaForm.Qualified)]
        public string FullColumns { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Excel, Form = XmlSchemaForm.Qualified)]
        public string FullRows { get; set; }

        [XmlAttribute(Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public string DefaultRowHeight { get; set; }

        [XmlElement("Column", Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public List<Column> Columns { get; set; }

        [XmlElement("Row", Namespace = XmlNamespaces.Spreadsheet, Form = XmlSchemaForm.Qualified)]
        public List<Row> Rows { get; set; }

        [XmlNamespaceDeclarations]
        public XmlSerializerNamespaces xmlns = new XmlSerializerNamespaces();

        public Table()
        {
            xmlns.Add("x", XmlNamespaces.Excel);
            xmlns.Add("ss", XmlNamespaces.Spreadsheet);
        }

        /// <summary>
        /// Create a System.Data.DataTable object from the worksheet's tabular data.
        /// </summary>
        public DataTable AsDataTable()
        {
            var result = new DataTable();
            // this method assumes the worksheet is arranged in tabular format
            // with header names on the first row
            foreach (var heading in from cell in Rows[0].Cells select cell.Data) {
                if (heading.Type.Equals("Number")) {
                    if (heading.Value.Contains(".")) {
                        result.Columns.Add(heading.Value, typeof(decimal));
                    } else {
                        result.Columns.Add(heading.Value, typeof(int));
                    }
                } else if (heading.Type.Equals("Boolean")) {
                    result.Columns.Add(heading.Value, typeof(bool));
                } else if (heading.Type.Equals("DateTime")) {
                    result.Columns.Add(heading.Value, typeof(DateTime));
                } else {
                    result.Columns.Add(heading.Value, typeof(string));
                }
            }
            // add the remaining rows
            foreach (var cells in from row in Rows.Skip(1) select row.Cells) {
                var values = new List<string>();
                int index = 1;
                // account for ss:Index and "skipped" columns
                foreach (var cell in cells) {
                    for (; !string.IsNullOrEmpty(cell.Index) && index < Convert.ToInt32(cell.Index); index++) {
                        values.Add("");
                    }
                    values.Add(cell.Data.Value);
                    if (string.IsNullOrEmpty(cell.Index)) index++;
                }
                result.Rows.Add(values.ToArray());
            }
            return result;
        }

        /// <summary>
        /// Generate Worksheet.Table object from System.Data.DataTable
        /// </summary>
        public static Table FromDataTable(DataTable source)
        {
            var result = new Table() {
                ExpandedColumnCount = $"{source.Columns.Count}",
                ExpandedRowCount = $"{source.Rows.Count + 1}",
                FullColumns = "1",
                FullRows = "1",
                Rows = new List<Row>(from DataRow dr in source.Rows
                                     select new Row {
                                         Cells = new List<Cell>(from obj in dr.ItemArray select new Cell { Data = CreateDataElement(obj) })
                                     }),
            };
            result.Rows.Insert(0, new Row {
                Cells = new List<Cell>(from DataColumn dc in source.Columns
                                       select new Cell {
                                           Data = new Data { Type = "String", Value = dc.ColumnName }
                                       })
            });
            return result;
        }

        private static Data CreateDataElement(object value)
        {
            string typeId;
            if (value is int || value is long || value is decimal || value is double || value is float) {
                typeId = "Number";
            } else if (value is bool) {
                typeId = "Boolean";
            } else if (value is DateTime) {
                typeId = "DateTime";
            } else {
                typeId = "String";
            }
            return new Data {
                Type = typeId,
                Value = Convert.ToString(value)
            };
        }
    }
}
