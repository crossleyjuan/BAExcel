using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BAExcel
{
    public class CellPos {
        public int Row { get; set; }
        public int Col { get; set; }

        private CellPos() { }

        public static CellPos CreateCellPos(int row, int col)
        {
            return new CellPos {
                Row = row,
                Col = col
            };
        }
    }

    public class CellRange
    {
        public CellPos StartRange { get; set; }

        public CellPos EndRange { get; set; }

        public static CellRange CreateCellRange(CellPos start, CellPos end)
        {
            return new CellRange
            {
                StartRange = start,
                EndRange = end
            };
        }
    }
}
