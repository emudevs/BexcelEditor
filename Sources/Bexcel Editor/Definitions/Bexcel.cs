using System.Collections.Generic;

namespace Bexcel_Editor.Definitions
{
    internal class Bexcel
    {
        public class Column
        {
            public string Name { get; set; }
        }

        public class Cell
        {
            public int Index { get; set; }
            public string Name { get; set; }
        }

        public class Row
        {
            public int Index1 { get; set; }
            public int Index2 { get; set; }

            public List<Cell> Cells { get; set; } = new List<Cell>();
        }

        public struct Unknown
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public int Number { get; set; }
        }

        public class Sheet
        {
            public string Name { get; set; }
            public int Type { get; set; }
            public int Index1 { get; set; }
            public int Index2 { get; set; }

            public List<Column> Columns { get; set; } = new List<Column>();
            public List<Row> Rows { get; set; } = new List<Row>();

            public List<Unknown> Unknown1 { get; set; } = new List<Unknown>();
            public List<Unknown> TableDetails { get; set; } = new List<Unknown>();

            private System.Data.DataTable _dt;
            public System.Data.DataTable Dt
            {
                get
                {
                    if (_dt == null)
                    {
                        _dt = BexcelConverter.ToDataTable(this); string s = "test";
                    }

                    return _dt;
                }
            }
        }

        public List<Sheet> Sheets { get; set; } = new List<Sheet>();
        public string FileEnding { get; set; }
    }
}