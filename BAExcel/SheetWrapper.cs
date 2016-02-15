using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace BAExcel
{
    public class SheetWrapper
    {
        HSSFSheet _sheet;
        BookWrapper _parent;

        public SheetWrapper(BookWrapper parent, HSSFSheet sheet)
        {
            _sheet = sheet;
            _parent = parent;
        }

        public string Name
        {
            get
            {
                return _sheet.SheetName;
            }
        }

        public object Get(CellRange range)
        {
            ICell cell = _sheet.GetRow(0).GetCell(0);
            return cell.StringCellValue;
        }

        private ICell GetCell(CellPos pos, bool create)
        {
            IRow row = _sheet.GetRow(pos.Row);
            if (row == null)
            {
                if (create == true)
                {
                    row = _sheet.CreateRow(pos.Row);
                }
                else
                {
                    return null;
                }
            }
            ICell cell = row.GetCell(pos.Col);
            if (cell == null)
            {
                if (create == true)
                {
                    cell = row.CreateCell(pos.Col);
                }
                else
                {
                    return null;
                }
            }
            return cell;
        }

        struct TextPosition
        {
            public int start;
            public int end;
        };

        public string GetText(CellPos pos)
        {
            ICell cell = GetCell(pos, false);
            if (cell != null)
            {
                return cell.StringCellValue;
            }
            else
            {
                return null;
            }
        }

        public void SetText(CellPos pos, string text)
        {
            List<TextPosition> positions = new List<TextPosition>();

            while (text.IndexOf("<b>") > -1)
            {
                int start = text.IndexOf("<b>");

                text = text.Substring(0, start) + text.Substring(start + 3);
                int end = text.IndexOf("</b>");
                text = text.Substring(0, end) + text.Substring(end + 4);
                positions.Add(new TextPosition { start = start, end = end });
            }
            HSSFRichTextString richtext = new HSSFRichTextString(text);
            IFont font1 = _parent.Internal.CreateFont();
            font1.Boldweight = (short)FontBoldWeight.Bold;
            foreach (TextPosition position in positions)
            {
                richtext.ApplyFont(position.start, position.end, font1);
            }

            ICell cell = GetCell(pos, true);
            cell.SetCellValue(richtext);
        }

        public void ApplyStyle(CellPos pos, string name)
        {
             ICellStyle style = _parent.GetStyle(name);

             ICell cell = GetCell(pos, false);
             if (cell != null)
             {
                 cell.CellStyle = style;
             }
        }

        public void SetColumnWidth(int col, int val)
        {
            _sheet.SetColumnWidth(col, val * 256);
        }
    }
}
