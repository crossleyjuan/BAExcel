using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;

namespace BAExcel
{
    public class BookWrapper
    {
        private HSSFWorkbook _book;
        private Dictionary<string, ICellStyle> _styles = new Dictionary<string, ICellStyle>();

        public BookWrapper(HSSFWorkbook book)
        {
            _book = book;
        }

        internal HSSFWorkbook Internal
        {
            get { return _book; }
        }

        public SheetWrapper CreateSheet(string name)
        {
            HSSFSheet sheet = (HSSFSheet)_book.CreateSheet(name);

            SheetWrapper sheetResult = new SheetWrapper(this, sheet);
            sheet.AlternativeFormula = false;
            sheet.AlternativeExpression = false;


            return sheetResult;
        }

        public void CreateStyle(string name, bool bold, short fontSize, bool wrapText, short indention)
        {
            //font style1: underlined, italic, red color, fontsize=20
            IFont font1 = _book.CreateFont();
//font1.Color = HSSFColor.Red.Index;
            font1.Boldweight = bold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal;
            //font1.IsItalic = true;
            //font1.Underline = FontUnderlineType.Double;
            font1.FontHeightInPoints = fontSize;

            //bind font with style 1
            ICellStyle style1 = _book.CreateCellStyle();
            style1.WrapText = wrapText;
            style1.SetFont(font1);
            style1.Alignment = HorizontalAlignment.Left;
            style1.VerticalAlignment = VerticalAlignment.Top;
            style1.Indention = indention;
            _styles.Add(name, style1);

        }

        internal ICellStyle GetStyle(string name)
        {
            return _styles[name];
        }

        public IEnumerator<SheetWrapper> Sheets() {
            return new SheetWrapperEnumerator<SheetWrapper>(this);
        }

        public class SheetWrapperEnumerator<T> : IEnumerator<SheetWrapper>
        {
            BookWrapper _book;
            System.Collections.IEnumerator _internalEnumerator;

            public SheetWrapperEnumerator(BookWrapper book)
            {
                _book = book;
                _internalEnumerator = book._book.GetEnumerator();
            }


            public SheetWrapper Current
            {
                get { return new SheetWrapper(_book, (HSSFSheet)_internalEnumerator.Current); }
            }

            public void Dispose()
            {
            }

            object System.Collections.IEnumerator.Current
            {
                get { return new SheetWrapper(_book, (HSSFSheet)_internalEnumerator.Current); }
            }

            public bool MoveNext()
            {
                return _internalEnumerator.MoveNext();
            }

            public void Reset()
            {
                _internalEnumerator.Reset();
            }
        }

        public void Save(string fileName)
        {
           FileStream stream = new FileStream(fileName, FileMode.Create);
            _book.Write(stream);
            stream.Close();
        }

        public sbyte[] GetBytes()
        {
            MemoryStream ms = new MemoryStream();
            _book.Write(ms);
            ms.Flush();
            byte[] result = ms.ToArray();
            ms.Close();

            sbyte[] sresult = new sbyte[result.Length];
            Buffer.BlockCopy(result, 0, sresult, 0, result.Length);
            return sresult;
        }
    }
}
