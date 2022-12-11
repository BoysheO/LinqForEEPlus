using System;
using System.Runtime.Serialization;

namespace LinqForEEPlus
{
    [DataContract]
    public readonly struct Cell {
        public readonly int   RowNum;
        public readonly int   ColNum;
        public readonly Sheet Sheet;

        public bool   IsLoaded   => Sheet.IsLoaded;
        public Cell   ColHead    => new Cell(Sheet, 1, ColNum);
        public string SheetName  => Sheet.Name;
        public Row    Row        => Sheet[RowNum];
        public object Value      => Sheet.ExcelWorksheet.Cells[RowNum, ColNum].Value;
        public bool   IsFlo      => StrValue != null && float.TryParse(StrValue, out _);
        public float  Flo        => float.Parse(StrValue);
        /// <summary>
        /// 保证cell已经加载而且不会是null/空白字符
        /// </summary>
        public bool   IsValuable => IsLoaded && !string.IsNullOrWhiteSpace(StrValue);

        /// <summary>
        /// 要么返回null，要么返回非空字段且前后不会是空白字符
        /// </summary>
        public string StrValue {
            get {
                var str = Value?.ToString();
                if (string.IsNullOrWhiteSpace(str)) return null; //这里就保证了Trim后不会为""
                return str.Trim();                               //此时必不为""
            }
        }
        
        [DataMember]
        public string DebugString {
            get => $"[{StrValue},{RowNum}行,{ColNum}|{Sheet.GetColumnLetter(ColNum)}列]";
            set { throw new NotSupportedException(); } //为了方便xml序列化
        }

        public Cell(Sheet sheet, int row, int col)
        {
            if (!sheet.IsLoaded) throw new ArgumentException("sheet is not loaded");
            Sheet  = sheet;
            RowNum = row;
            ColNum = col;
        }

        public override string ToString()
        {
            return IsLoaded?( Value == null ? "(nil)" : Value.ToString()):("noload");
        }

        public bool TryGetFlo(out float flo)
        {
            return float.TryParse(StrValue, out flo);
        }

        public bool IsEquals(Cell cell)
        {
            return cell.Sheet.IsEquals(this.Sheet) && cell.RowNum == this.RowNum && cell.ColNum == this.ColNum;
        }

        public static implicit operator string(Cell cell)
        {
            return cell.StrValue;
        }

    }
}