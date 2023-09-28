using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel; // add Microsoft Excel COM reference
using DT = System.Data;

namespace RpaLib.Tracing
{
    public enum InsertMethod
    {
        AsRow,
        AsColumn
    }
    public class Excel
    {
        public Application Application { get; }
        public Workbook Workbook { get; private set; }
        public string FullFilePath { get; private set; }
        public Worksheet Worksheet { get; private set; }
        public string SheetName { get; private set; }
        public Dictionary<string, int> UsedRangeCount { get; private set; }

        public Excel(string filePath, string sheetName)
        {
            Application = new Application();
            Application.DisplayAlerts = false; // disable pop-ups when overwriting

            FullFilePath = Path.GetFullPath(filePath);
            SheetName = sheetName;

            Workbook = Application.Workbooks.Open(FullFilePath);
            Worksheet = Workbook.Sheets.Item[SheetName];

            UpdateUsedRangeCount();
        }

        private void UpdateUsedRangeCount()
        {
            UsedRangeCount = new Dictionary<string, int>()
            {
                { "rows", Worksheet.UsedRange.Rows.Count },
                { "cols", Worksheet.UsedRange.Columns.Count }
            };
        }

        public static int[] Range(int start, int end)
        {
            List<int> range = new List<int>();

            for (int i = start; i <= end; i++)
                range.Add(i);

            return range.ToArray();
        }

        public void Quit()
        {
            Workbook.Close();
            Application.Quit();
        }

        public void Save(string path = null)
        {
            if (path == null)
                Workbook.Save();
            else
                Workbook.SaveAs(Path.GetFullPath(path));
        }

        public Dictionary<string, string> GetActiveSheet()
        {
            return new Dictionary<string, string>()
            {
                { "name", Workbook.ActiveSheet.Name },
                { "index", Workbook.ActiveSheet.Index }
            };
        }

        public void SetActiveSheet(dynamic nameOrIndex)
        {
            if (nameOrIndex is string || nameOrIndex is int)
                _ = Workbook.Sheets.Item[nameOrIndex];
            else
                throw new ArgumentException("SetActiveSheet() receives nameOrIndex of types int or string only.");
        }

        public void ToggleVisible()
        {
            Application.Visible = !Application.Visible;
        }

        public void ToggleVisible(bool makeVisible)
        {
            if (makeVisible)
                Application.Visible = true;
            else
                Application.Visible = false;
        }

        public string ReadCellTest(int row, int col)
        {
            return (string)Application.Cells[row, col].Text;
        }

        public string ReadCell(int row, int col)
        {
            return (string)Worksheet.Rows.Item[row].Columns.Item[col].Text;
        }

        public string[] ReadCell(int row, int[] cols)
        {
            List<string> cells = new List<string>();

            foreach (int col in cols)
                cells.Add((string)Worksheet.Rows.Item[row].Columns.Item[col].Text);

            return cells.ToArray();
        }

        public string[] ReadCell(int[] rows, int col)
        {
            List<string> cells = new List<string>();

            foreach (int row in rows)
                cells.Add((string)Worksheet.Rows.Item[row].Columns.Item[col].Text);

            return cells.ToArray();
        }

        public string[][] ReadCell(int[] rows, int[] cols)
        {
            List<string[]> table = new List<string[]>();

            foreach (int row in rows)
            {
                List<string> column = new List<string>();
                foreach (int col in cols)
                {
                    column.Add((string)Worksheet.Rows.Item[row].Columns.Item[col].Text);
                }
                table.Add(column.ToArray());
            }
            return table.ToArray();
        }

        public string[][] ReadCell(int startRow, int startCol, int endRow, int endCol)
        {
            return ReadCell(Range(startRow, endRow), Range(startCol, endCol));
        }

        public string[] ReadCell(int row, int startCol, int endCol)
        {
            return ReadCell(row, Range(startCol, endCol));
        }

        public void WriteCell(int row, int col, string value)
        {
            Worksheet.Rows.Item[row].Columns.Item[col] = value;
            UpdateUsedRangeCount();
        }

        public void WriteCell(int row, int col, string[] values, InsertMethod rowOrCol)
        {
            for (int j = 0; j < values.Count(); j++)
            {
                switch (rowOrCol)
                {
                    case InsertMethod.AsRow:
                        Worksheet.Rows.Item[row].Columns[col + j] = values[j];
                        break;
                    case InsertMethod.AsColumn:
                        Worksheet.Rows.Item[row + j].Columns[col] = values[j];
                        break;
                    default:
                        throw new ArgumentException($"Invalid insertion method: {rowOrCol}");
                }
            }
            UpdateUsedRangeCount();
        }

        public void WriteCell(int row, int col, string[][] tableOfValues)
        {
            for (int i = 0; i < tableOfValues.Count(); i++)
                WriteCell(i + row, col, tableOfValues[i], InsertMethod.AsRow);
        }

        public static dynamic Helper(
            string file,
            string worksheet,
            bool visible = false,
            string saveAs = null,
            string[] writeRow = null,
            string[] writeCol = null,
            string[][] writeTable = null,
            DT.DataTable writeDataTable = null,
            bool readAll = false,
            int readRow = -1,
            int readCol = -1,
            int startRow = 1,
            int startCol = 1,
            int endRow = -1,
            int endCol = -1,
            bool appendRow = false,
            bool appendCol = false)
        {
            dynamic returnValue = null;
            string fileFullPath = Path.GetFullPath(file);
            Excel excel = new Excel(fileFullPath, worksheet);

            excel.ToggleVisible(makeVisible: visible);

            if (appendRow)
                startRow = excel.UsedRangeCount["rows"] + 1;

            if (appendCol)
                startCol = excel.UsedRangeCount["cols"] + 1;

            if (endRow < 1)
                endRow = excel.UsedRangeCount["rows"];

            if (endCol < 1)
                endCol = excel.UsedRangeCount["cols"];

            if (writeRow != null)
                excel.WriteCell(startRow, startCol, writeRow, InsertMethod.AsRow);
            else if (writeCol != null)
                excel.WriteCell(startRow, startCol, writeCol, InsertMethod.AsColumn);
            else if (writeTable != null)
                excel.WriteCell(startRow, startCol, writeTable);
            else if (writeDataTable != null)
            {
                List<string[]> table = new List<string[]>();
                for (int i = 0; i < writeDataTable.Rows.Count; i++)
                {
                    List<string> row = new List<string>();
                    for (int j = 0; j < writeDataTable.Columns.Count; j++)
                    {
                        row.Add($"{writeDataTable.Rows[i][j]}");
                    }
                    table.Add(row.ToArray());
                }
                excel.WriteCell(startRow, startCol, table.ToArray());
            }

            if (readAll)
                returnValue =
                    excel.ReadCell(startRow, startCol,
                        excel.UsedRangeCount["rows"],
                        excel.UsedRangeCount["cols"]);
            else if (readRow > 0)
                returnValue =
                    excel.ReadCell(readRow, Excel.Range(startCol, endCol));
            else if (readCol > 0)
                returnValue =
                    excel.ReadCell(Excel.Range(startRow, endRow), readCol);

            excel.Save(saveAs);
            excel.Quit();

            return returnValue;
        }
    }
}
