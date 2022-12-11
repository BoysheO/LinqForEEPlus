using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace LinqForEEPlus
{
    /// <summary>
    /// 工作薄
    /// </summary>
    public readonly struct Book : IEnumerable<Sheet> {
        /// <summary>
        /// 是否已经通过构造函数构造
        /// </summary>
        public bool IsLoaded => ExcelWorkbook != null;

        public readonly ExcelWorkbook ExcelWorkbook;
        public readonly string        FullFileName;
        public readonly string        ShortFileName;

        /// <summary>
        /// 表的数量
        /// </summary>
        public int Count => ExcelWorkbook.Worksheets.Count;

        public Book(ExcelWorkbook book, string fileFullName)
        {
            ExcelWorkbook = book ?? throw new ArgumentNullException(nameof(book));
            FullFileName  = fileFullName;
            ShortFileName = Path.GetFileName(fileFullName);
        }

        /// <summary>
        /// 从0开始算起的数组标号
        /// </summary>
        public Sheet this[int index] => new Sheet(ExcelWorkbook.Worksheets[index]);

        /// <summary>
        /// 不保证其正确性 
        /// </summary>
        public Sheet this[string name] => new Sheet(ExcelWorkbook.Worksheets[name]);

        public IEnumerator<Sheet> GetEnumerator()
        {
            return new BookEnumerator(this);
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public override string ToString()
        {
            if (!IsLoaded) return "(noload)";
            return string.IsNullOrWhiteSpace(FullFileName) ? this.GetType().Name : FullFileName;
        }

        private struct BookEnumerator : IEnumerator<Sheet> {
            private readonly Book book;
            private          int  cur;
            private          int  maxcount;

            public BookEnumerator(Book book)
            {
                this.book = book;
                cur       = -1;
                maxcount  = book.Count;
            }

            public bool MoveNext()
            {
                cur++;
                return cur < maxcount;
            }

            public void Reset()
            {
                cur = -1;
            }

            public Sheet Current => book[cur];

            object IEnumerator.Current => Current;

            public void Dispose()
            {
                Reset();
            }
        }
    }
}