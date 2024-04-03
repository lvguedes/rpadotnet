using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel; // add Microsoft Excel COM reference
using Microsoft.Office.Core;
using DataTable = System.Data.DataTable;
using System.Data;
using RpaLib.ProcessAutomation;
using RpaLib.Tracing.Exceptions;
using System.Runtime.InteropServices;
using RpaLib.Exceptions;

namespace RpaLib.Tracing
{
    public class Excel
    {
        public const string DefaultSheetName = "Sheet1";
        public Application Application { get; }
        public Workbook Workbook { get; private set; }
        public string FullFilePath { get; private set; }
        public Worksheet Worksheet { get; private set; }
        public string SheetName { get; private set; }
        public Dictionary<string, int> UsedRangeCount { get; private set; }

        public Excel(string filePath, string sheetName, bool disableMacros = false)
        {
            Application = new Application();
            Application.DisplayAlerts = false; // disable pop-ups when overwriting

            if (disableMacros) Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            FullFilePath = Ut.GetFullPath(filePath);
            SheetName = sheetName;

            Workbook = Application.Workbooks.Open(FullFilePath, ReadOnly: disableMacros);
            try
            {
                Worksheet = Workbook.Sheets.Item[SheetName];
            }
            catch (COMException ex)
            {
                if (Ut.IsMatch(ex.Message, @"Invalid index\. \(Exception from HRESULT: 0x8002000B \(DISP_E_BADINDEX\)\)"))
                {
                    throw new WorksheetNotFoundException(sheetName, filePath);
                }
                else
                {
                    throw ex;
                }
            }
            

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
            Workbook.Close(SaveChanges: false);
            Application.Quit();
        }

        public void Save(string path = null)
        {
            if (path == null)
                Workbook.Save();
            else
                Workbook.SaveAs(Ut.GetFullPath(path));
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
                throw new RpaLibArgumentException("SetActiveSheet() receives nameOrIndex of types int or string only.");
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

        public string[][] ReadCell(int[] rows, int[] cols, bool breakAtEmptyLine = false)
        {
            List<string[]> table = new List<string[]>();

            foreach (int row in rows)
            {
                List<string> line = new List<string>();
                foreach (int col in cols)
                {
                    line.Add((string)Worksheet.Rows.Item[row].Columns.Item[col].Text);
                }
                if (IsEmptyLine(line.ToArray()) && breakAtEmptyLine)
                    break;
                else
                    table.Add(line.ToArray());
            }
            return table.ToArray();
        }

        public string[][] ReadCell(int startRow, int startCol, int endRow, int endCol, bool breakAtEmptyLine = false)
        {
            return ReadCell(Range(startRow, endRow), Range(startCol, endCol), breakAtEmptyLine);
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
                        throw new RpaLibArgumentException($"Invalid insertion method: {rowOrCol}");
                }
            }
            UpdateUsedRangeCount();
        }

        public void WriteCell(int row, int col, string[][] tableOfValues)
        {
            for (int i = 0; i < tableOfValues.Count(); i++)
                WriteCell(i + row, col, tableOfValues[i], InsertMethod.AsRow);
        }

        // Finalizer method
        ~Excel()
        {
            Application.Quit();
        }

        public static dynamic Helper(
            string file,
            string worksheet,
            bool visible = false,
            string saveAs = null,
            string[] writeRow = null,
            string[] writeCol = null,
            string[][] writeTable = null,
            DataTable writeDataTable = null,
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
            string fileFullPath = Ut.GetFullPath(file);
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

        public static DataTable ReadAll(string filePath, string sheetName, bool visible = false, bool disableMacros = false)
        {
            string fileFullPath = Ut.GetFullPath(filePath);
            Excel excel = new Excel(fileFullPath, sheetName, disableMacros);

            excel.ToggleVisible(visible);

            int startRow = 1;
            int startCol = 1;

            var contents = excel.ReadCell(startRow, startCol,
                            excel.UsedRangeCount["rows"],
                            excel.UsedRangeCount["cols"],
                            breakAtEmptyLine: true);

            excel.Quit();

            return ArrayStrTableToDataTable(contents);
        }

        public static void WriteNextFreeRow(string filePath, DataTable table, string sheetName = DefaultSheetName, bool visible = false, bool disableMacros = false)
        {
            string fileFullPath = Ut.GetFullPath(filePath);
            Excel excel = new Excel(fileFullPath, sheetName, disableMacros);

            excel.ToggleVisible(visible);

            int startRow = excel.UsedRangeCount["rows"] + 1;
            int startCol = 1;

            List<string[]> rowList = new List<string[]>();
            for (int i = 0; i < table.Rows.Count; i++)
            {
                List<string> row = new List<string>();
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    row.Add($"{table.Rows[i][j]}");
                }
                rowList.Add(row.ToArray());
            }
            excel.WriteCell(startRow, startCol, rowList.ToArray());

            excel.Save();
            excel.Quit();
        }

        public static string[][] Transpose(string[][] sheetArray)
        {
            List<List<string>> columns = new List<List<string>>();

            foreach (string col in sheetArray[0])
            {
                var column = new List<string>();
                columns.Add(column);
            }

            foreach (var row in sheetArray)
            {
                for (int i = 0; i < sheetArray[0].GetLength(0); i++)
                {
                    columns[i].Add(row[i]);
                }
            }

            return ConvList2Array(columns);
        }

        private static string[][] ConvList2Array(List<List<string>> list)
        {
            List<string[]> wrapper = new List<string[]>();
            
            foreach (var item in list)
            {
                wrapper.Add(item.ToArray());
            }

            return wrapper.ToArray();
        }

        public static bool IsEmptyLine(string[] line)
        {
            bool isEmpty = true;
            foreach (var col in line)
            {
                if (!string.IsNullOrEmpty(col) && !Ut.IsMatch(col, @"^\s+$"))
                {
                    isEmpty = false;
                    break;
                }
            }
            return isEmpty;
        }

        // converts an 2 dimensional array of strings (aka String Table) to a DataTable
        public static DataTable ArrayStrTableToDataTable(string[][] arrayStrTable)
        {
            DataTable dataTable = new DataTable();

            // Define dataColumns according to arrayStrTable header
            var headerCols = arrayStrTable[0];
            foreach (var col in headerCols)
            {
                var dataColumn = new DataColumn();
                dataColumn.ColumnName = col;
                dataColumn.DataType = typeof(string);
                dataTable.Columns.Add(dataColumn);
            }

            // Define dataRows from definition of dataColumn and assign arrayStrTable values to it
            var listStrTable = new List<string[]>(arrayStrTable);
            listStrTable.RemoveAt(0); // remove the header row
            foreach (var row in listStrTable)
            {
                DataRow dataRow = dataTable.NewRow();
                for (var j = 0; j < row.Length; j++)
                {
                    var currentColName = headerCols[j];
                    var currentColIndex = j;
                    dataRow[j] = row[j];
                }
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }
    }
}


/* Common Interop Error that might occur again once in while
 * 
 * System.InvalidCastException HResult=0x80004002 
 *   Message=Unable to cast COM object of type 'Microsoft.Office.Interop.Excel.ApplicationClass' to interface type 'Microsoft.Office.Interop.Excel._Application'.
 *   This operation failed because the QueryInterface call on the COM component for the interface with IID '{000208D5-0000-0000-C000-000000000046}' failed due
 *   to the following error: 
 *     Interface not registered (Exception from HRESULT: 0x80040155). 
 *     Source=mscorlib StackTrace: 
 *       at System.StubHelpers.StubHelpers.GetCOMIPFromRCW(Object objSrc, IntPtr pCPCMD, IntPtr& ppTarget, Boolean& pfNeedsRelease) 
 *       at Microsoft.Office.Interop.Excel.ApplicationClass.set_DisplayAlerts(Boolean RHS) 
 *       at RpaLib.Tracing.Excel..ctor(String filePath, String sheetName) in C:\lib_project\Tracing\Excel.cs:line 31 
 *       at RpaLib.Tracing.Excel.ReadAll(String filePath, String sheetName) in C:\lib_project\Tracing\Excel.cs:line 271 
 *       at AlteracaoEstrutura.Program.Main(String[] args) in C:\main_project\Program.cs:line 19
 *
 *  Whenever it occur again, REPAIR OFFICE to fix this issue.
 */