using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Tracing
{
    public class RangeCoords
    {
        public CellCoords First { get; set; } = new CellCoords();
        public CellCoords Last { get; set; } = new CellCoords();

        public RangeCoords() { }
        public RangeCoords(CellCoords first, CellCoords last)
        {
            First = first;
            Last = last;
        }
        public RangeCoords(int firstRow, ExcelColumn firstCol, int lastRow, ExcelColumn lastCol)
        {
            First = new CellCoords(firstRow, firstCol);
            Last = new CellCoords(lastRow, lastCol);
        }
    }
}
