using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Tracing
{
    public class CellCoords
    {
        public int Row { get; set; } = 0;
        public ExcelColumn Col { get; set; } = ExcelColumn.None;

        public CellCoords() { }
        public CellCoords(int row, ExcelColumn col)
        {
            Row = row;
            Col = col;
        }
        public CellCoords(Range range) 
        : this(range.Row, (ExcelColumn)range.Column) { }
    }
}
