using System.Collections;
using System.Collections.Generic;

// ReSharper disable All

namespace LinqForEEPlus
{
    public readonly struct Row : IEnumerable<Cell> {
        public readonly Sheet Sheet;
        public readonly int   RowNum;
        public          bool  IsLoaded  => Sheet.IsLoaded;
        public          Cell  FirstCell => new Cell(Sheet, RowNum, 1);
        public          Cell? CHKEYCell => this["CHKEY"];
        public          Cell? ENKEY     => this["ENKEY"];

        public bool IsEmptyCol {
            get {
                foreach (var cell in this) {
                    if (cell.StrValue != null) return false;
                }

                return true;
            }
        }

        public Cell this[int col] => new Cell(Sheet, RowNum, col);

        public Cell? this[string colhead] {
            get {
                for (int col = 1; col <= Sheet.ColMax; col++) {
                    var head = Sheet[1, col].ToString();
                    if (head == colhead) {
                        return new Cell(Sheet, RowNum, col);
                    }
                }

                return null;
            }
        }

        public Row(Sheet mSheet, int rowNum)
        {
            this.Sheet  = mSheet;
            this.RowNum = rowNum;
            if (RowNum < 1) throw new System.ArgumentOutOfRangeException(nameof(rowNum));
        }

        public IEnumerator<Cell> GetEnumerator()
        {
            return new RowEnumerator(Sheet, RowNum);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public bool IsEquals(Row row)
        {
            if (row.Sheet.IsEquals(this.Sheet) && row.RowNum == this.RowNum) return true;
            return false;
        }

        public override string ToString()
        {
            return IsLoaded ? $"{this[1]}|[{this[2]}|[{this[3]}]|..." : "(noload)";
        }

        #region RowEnumerator

        private struct RowEnumerator : IEnumerator<Cell> {
            private Sheet Sheet;
            private int   row;
            private int   colmax;
            private int   col;

            public RowEnumerator(Sheet sheet, int row)
            {
                Sheet    = sheet;
                this.row = row;
                colmax   = sheet.ColMax;
                col      = 0;
            }

            public Cell Current => new Cell(Sheet, row, col);

            object IEnumerator.Current => Current;

            public void Dispose() { }

            public bool MoveNext()
            {
                col++;
                return col <= colmax;
            }

            public void Reset()
            {
                col = 0;
            }
        }

        #endregion
    }
}