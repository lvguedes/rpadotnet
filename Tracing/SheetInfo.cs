using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Exceptions;

namespace RpaLib.Tracing
{
    public class SheetInfo
    {
        public string Name { get; set; } = Excel.DefaultSheetName;

        private int _firstRow = 0;
        public int FirstRow
        {
            get => _firstRow;
            set
            {
                ParseExcelCellIndex(value);
                _firstRow = value;
            }
        }

        private ExcelColumn _firstCol = ExcelColumn.None;
        public ExcelColumn FirstCol
        {
            get => _firstCol;
            set
            {
                ParseExcelCellIndex(value);
                _firstCol = value;
            }
        }

        private int _lastRow = 0;
        public int LastRow
        {
            get => _lastRow;
            set
            {
                ParseExcelCellIndex(value);
                _lastRow = value;
            }
        }

        private ExcelColumn _lastCol = ExcelColumn.None;
        public ExcelColumn LastCol
        {
            get => _lastCol;
            set
            {
                ParseExcelCellIndex(value);
                _lastCol = value;
            }
        }

        private void ParseExcelCellIndex(int value)
        {
            if (value <= 0)
                throw new RpaLibArgumentException($"Excel cell index must be integer greater than 0. Entered: {value}");
        }

        private void ParseExcelCellIndex(ExcelColumn excelColumn)
        {
            ParseExcelCellIndex((int)excelColumn);
        }
    }
}
