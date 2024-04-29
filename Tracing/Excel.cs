using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading;
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
        public const string ProcessName = "EXCEL";
        public Application Application { get; }
        public Workbook Workbook { get; private set; }
        public string FullFilePath { get; private set; }
        public Worksheet Worksheet { get; private set; }
        public string SheetName { get; private set; }
        public SheetUsedRange UsedRangeCount { get; private set; }

        public Excel(string filePath, string sheetName, bool disableMacros = false)
        {
            Application = new Application();
            Application.DisplayAlerts = false; // disable pop-ups when overwriting
            var disableMacrosOnlyXlsm = disableMacros && Ut.IsMatch(filePath, @"\.xlsm$");

            if (disableMacrosOnlyXlsm) Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            FullFilePath = Ut.GetFullPath(filePath);
            SheetName = sheetName;

            Workbook = Application.Workbooks.Open(FullFilePath, ReadOnly: disableMacrosOnlyXlsm);
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

            UsedRangeCount = new SheetUsedRange();
            UpdateUsedRangeCount();
        }

        private void UpdateUsedRangeCount()
        {
            UsedRangeCount.Rows = Worksheet.UsedRange.Rows.Count;
            UsedRangeCount.Cols = Worksheet.UsedRange.Columns.Count;
        }

        public static int[] Range(int start, int end)
        {
            List<int> range = new List<int>();

            for (int i = start; i <= end; i++)
                range.Add(i);

            return range.ToArray();
        }

        public Range GetRange(string excelRange)
        {
            var range = Worksheet.Range[excelRange];
            return range;
        }

        public Range GetRange(string firstCell, string lastCell)
        {
            var range = Worksheet.Range[firstCell, lastCell];
            return range;
        }

        public Range GetRange(int firstCellRow, int firstCellCol, int lastCellRow, int lastCellCol)
        {
            var range = Worksheet.Range[Worksheet.Cells[firstCellRow, firstCellCol], Worksheet.Cells[lastCellRow, lastCellCol]];
            return range;
        }

        // usage: InsertFormula("=VLOOKUP()", "A1:B20");
        public void InsertFormula(string formula, string excelRange)
        {
            Range formulaFullRange = GetRange(excelRange);
            //Range formulaCell = formulaFullRange.Cells[1];
            //formulaCell.Formula = formula;
            //formulaCell.AutoFill(formulaFullRange, XlAutoFillType.xlFillDefault);
            InsertFormula(formulaFullRange, formula);
        }

        // usage: InsertFormula("=VLOOKUP()", "A1", "B20");
        public void InsertFormula(string formula, string firstCell, string lastCell)
        {
            Range formulaFullRange = GetRange(firstCell, lastCell);
            InsertFormula(formulaFullRange, formula);
        }

        // To work with sheet used ranges.
        // usage:
        //    InsertFormula("=VLOOKUP()", 1, 1, 20, 2); // insert formula in range "A1:B20"
        //    InsertFormula("=VLOOKUP()", 1, 1); //last cell will be the last used firstCellRow and column
        //    InsertFormula("=VLOOKUP()", 1, 1, 20); // last cell column will be the last used
        //    InsertFormula("=VLOOKUP()", 1, 1, lastCellCol: 2); // last cell firstCellRow will be the last used
        public void InsertFormula(string formula, int firstCellRow, int firstCellCol, int lastCellRow = -1, int lastCellCol = -1)
        {
            if (lastCellRow < 0)
                lastCellRow = UsedRangeCount.Rows;

            if (lastCellCol < 0)
                lastCellCol = UsedRangeCount.Cols;

            Range formulaFullRange = GetRange(firstCellRow, firstCellCol, lastCellRow, lastCellCol);
            InsertFormula(formulaFullRange, formula);
        }

        private void InsertFormula(Range range, string excelFormula)
        {
            range.Formula = excelFormula;
        }

        public void Quit()
        {
            try
            {
                //Workbook.Close(SaveChanges: false);
                Application.DisplayAlerts = false; // disable confirmation popups
                Application.Quit(); // Maybe an interop bug: just puts in background instead of closing.

                /*
                // try quit while instance still alive
                while (true)
                {
                    Ut.CloseMainWindow(Application.Hwnd, backgroundKill: true);
                    Thread.Sleep(2000);
                    //Ut.KillProcess(Application.Hwnd);

                    if (CheckInstanceAlive())
                        Thread.Sleep(5000); // wait only if still alive
                    else
                        break;
                }
                */
            }
            catch
            {
                Ut.KillProcess(ProcessName);
            } 
               
        }

        // return true if instance alive and false otherwise
        private bool CheckInstanceAlive()
        {
            try
            {
                // try to access a COM obj property
                // to see if object is still alive
                _ = Application.Path;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void SaveAndQuit()
        {
            Save();
            Quit();
        }

        public void SaveAndQuit(string path)
        {
            Save(path);
            Quit();
        }

        public void Save()
        {
            Workbook.Save();
        }

        public void Save(string path)
        {
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

        public string[][] ReadCellMultiProc(int[] rows, int[] cols, bool breakAtEmptyLine = false, int threads = 8)
        {
            // divide the range into threads equally
            int rowNum = rows.Length / threads;
            var tasks = new List<Task<List<string[]>>>();

            for (int i = 1; i <= threads; i++)
            {
                int minRow = (i - 1) * rowNum + 1;
                int maxRow = i * rowNum;

                if (i == threads)
                    maxRow += (rows.Length % threads);

                var rowRange = Range(minRow, maxRow);

                tasks.Add(Task.Factory.StartNew(() => ReadCellEditAppendRows(rowRange, cols, breakAtEmptyLine))); 
            }

            Task.WaitAll(tasks.ToArray());

            List<string[]> table = null;

            for (int i = 0; i < tasks.Count; i++)
            {
                var task = tasks[i];

                if (table == null)
                    table = task.Result;
                else
                    task.Result.ForEach(x => table.Add(x));
            }

            return table.ToArray();
        }

        public List<string[]> ReadCellEditAppendRows(int[] rows, int[] cols, bool breakAtEmptyLine = false)
        {
            List<string[]> table = new List<string[]>();

            foreach (int row in rows)
            {
                List<string> line = new List<string>();
                foreach (int col in cols)
                {
                    while (true)
                    {
                        try
                        {
                            line.Add((string)Worksheet.Rows.Item[row].Columns.Item[col].Text);
                            break;
                        }
                        catch (COMException ex)
                        {
                            if (Ut.IsMatch(ex.Message, @"The message filter indicated that the application is busy\."))
                                continue;
                            else
                                throw ex;
                        }
                    }
                }
                if (IsEmptyLine(line.ToArray()) && breakAtEmptyLine)
                    break;
                else
                    table.Add(line.ToArray());
            }
            return table;
        }

        public string[][] ReadCell(int[] rows, int[] cols, bool breakAtEmptyLine = false)
        {
            return ReadCellEditAppendRows(rows, cols, breakAtEmptyLine).ToArray();
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

        public void WriteCell(int firstCellRow, int firstCellCol, string[] values, InsertMethod rowOrCol)
        {
            for (int j = 0; j < values.Count(); j++)
            {
                switch (rowOrCol)
                {
                    case InsertMethod.AsRow:
                        Worksheet.Rows.Item[firstCellRow].Columns[firstCellCol + j] = values[j];
                        break;
                    case InsertMethod.AsColumn:
                        Worksheet.Rows.Item[firstCellRow + j].Columns[firstCellCol] = values[j];
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
            Quit(); 
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
                startRow = excel.UsedRangeCount.Rows + 1;

            if (appendCol)
                startCol = excel.UsedRangeCount.Cols + 1;

            if (endRow < 1)
                endRow = excel.UsedRangeCount.Rows;

            if (endCol < 1)
                endCol = excel.UsedRangeCount.Cols;

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
                        excel.UsedRangeCount.Rows,
                        excel.UsedRangeCount.Cols);
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

        public static DataTable ReadAll(string filePath, string sheetName, bool visible = false, bool disableMacros = false, bool breakAtEmptyLine = true, int startRow = 1, int startCol = 1)
        {
            string fileFullPath = Ut.GetFullPath(filePath);
            bool disableMacroOnlyIfXlsm = disableMacros && Ut.IsMatch("filePath", @"\.xlsm$");
            Excel excel = new Excel(fileFullPath, sheetName, disableMacroOnlyIfXlsm);

            excel.ToggleVisible(visible);

            var contents = excel.ReadCell(startRow, startCol,
                            excel.UsedRangeCount.Rows,
                            excel.UsedRangeCount.Cols,
                            breakAtEmptyLine: true);

            //var contents = excel.ReadCellMultiProc(Range(startRow, excel.UsedRangeCount.Rows), Range(startCol, excel.UsedRangeCount.Cols), breakAtEmptyLine: breakAtEmptyLine);

            excel.Quit();

            return ArrayStrTableToDataTable(contents);
        }

        // Writes an array of values as a column or firstCellRow into the sheet. It starts at the cell specified by firstCellRow and firstCellCol values.
        // If firstCellRow or firstCellCol are negative, it will consider the next free one. 
        public static void WriteAll(string filePath, string sheetName, int row, int col, string[] values, InsertMethod rowOrCol, bool visible = false)
        {
            var excel = new Excel(filePath, sheetName, disableMacros: false);
            excel.ToggleVisible(visible);

            if (row < 0)
                row = excel.UsedRangeCount.Rows;

            if (col < 0)
                col = excel.UsedRangeCount.Cols;

            excel.WriteCell(row, col, values, rowOrCol);
            excel.Save();
            excel.Quit();
        }

        // TODO: code validation for excel Range (accept only: A1, A1:B2, A:A)
        // formula must be in the English like manner =VLOOKUP(G2,A:C,3,0) 
        public static void InsertFormula(string filePath, string sheetName, string excelRange, string formula, bool visible = false)
        {
            var excel = new Excel(filePath, sheetName, disableMacros: false);
            excel.ToggleVisible(visible);

            excel.InsertFormula(formula, excelRange);

            excel.SaveAndQuit();
        }

        public static void InsertFormula(string filePath, string sheetName, string formula, int firstCellRow, int firstCellCol, int lastCellRow = -1, int lastCellCol = -1, bool visible = false)
        {
            var excel = new Excel(filePath, sheetName, disableMacros: false);
            excel.ToggleVisible(visible);

            excel.InsertFormula(formula, firstCellRow, firstCellCol, lastCellRow, lastCellCol);

            excel.SaveAndQuit();
        }

        public static void InsertFormula(string filePath, string sheetName, string formula, string firstCell, string lastCell, bool visible = false)
        {
            var excel = new Excel(filePath, sheetName, disableMacros: false);
            excel.ToggleVisible(visible);

            excel.InsertFormula(formula, firstCell, lastCell);

            excel.SaveAndQuit();
        }

        public static void WriteNextFreeRow(string filePath, DataTable table, string sheetName = DefaultSheetName, bool visible = false, bool disableMacros = false)
        {
            string fileFullPath = Ut.GetFullPath(filePath);
            Excel excel = new Excel(fileFullPath, sheetName, disableMacros);

            excel.ToggleVisible(visible);

            int startRow = excel.UsedRangeCount.Rows + 1;
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
            listStrTable.RemoveAt(0); // remove the header firstCellRow
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