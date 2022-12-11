using OfficeOpenXml;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace LinqForEEPlus {
    public readonly struct Sheet : IEnumerable<Row> {
        public ExcelWorksheet ExcelWorksheet { get; }
        public bool           IsLoaded       => ExcelWorksheet != null;
        public string         Name           => ExcelWorksheet.Name;
        public Row            FirstRow       => this.FirstOrDefault();

        /// <summary>
        /// 有效列号最大值
        /// </summary>
        public readonly int ColMax;

        /// <summary>
        /// 有效行号最大值
        /// </summary>
        public readonly int RowMax;

        public Cell this[int row, int colum] => new Cell(this, row, colum);

        public Row this[int row] => new Row(this, row);

        public Sheet(ExcelWorksheet excelWorksheet)
        {
            ExcelWorksheet = excelWorksheet;
            ColMax         = excelWorksheet.Dimension.Columns;
            RowMax         = excelWorksheet.Dimension.Rows;
        }

        public static string GetColumnLetter(int iColumnNumber, bool fixedCol=false)
        {
            if (iColumnNumber < 1) {
                //throw new Exception("Column number is out of range");
                return "#REF!";
            }

            string sCol = "";
            do {
                sCol          = ((char) ('A' + (iColumnNumber - 1) % 26)).ToString() + sCol;
                iColumnNumber = (iColumnNumber - (iColumnNumber - 1) % 26) / 26;
            } while (iColumnNumber > 0);

            return fixedCol ? "$" + sCol : sCol;
        }

        public IEnumerator<Row> GetEnumerator()
        {
            return new SheetEnumerator(this);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IEnumerable<Row> GetEnumerable(int startRow)
        {
            return new Enumerable(this, startRow);
        }

        public bool IsEquals(Sheet sheet) => sheet.ExcelWorksheet == this.ExcelWorksheet;

        public override string ToString()
        {
            return IsLoaded?Name:"(noload)";
        }

        #region Enumerable

        private struct Enumerable : IEnumerable<Row> {
            private Sheet mSheet;
            private int   startRow;

            public Enumerable(Sheet sheet) : this(sheet, 1) { }

            public Enumerable(Sheet sheet, int startRow)
            {
                this.mSheet   = sheet;
                this.startRow = startRow;
            }

            public IEnumerator<Row> GetEnumerator()
            {
                return new SheetEnumerator(mSheet, startRow);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return new SheetEnumerator(mSheet, startRow);
            }
        }

        #endregion

        #region SheetEnumerator

        private struct SheetEnumerator : IEnumerator<Row> {
            private Sheet worksheet;
            private int   row;
            private int   StartRow;
            private int   rowmax;

            public SheetEnumerator(Sheet worksheet, int startRow = 1)
            {
                if (startRow < 1) throw new System.ArgumentOutOfRangeException();
                this.worksheet = worksheet;
                rowmax         = worksheet.RowMax;
                StartRow       = startRow - 1;
                row            = StartRow;
                Reset();
            }

            public Row Current => new Row(worksheet, row);

            object IEnumerator.Current => Current;

            public void Dispose() { }

            public bool MoveNext()
            {
                row++;
                return row <= rowmax;
            }

            public void Reset()
            {
                row = StartRow;
            }
        }

        #endregion
    }
}