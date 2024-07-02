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
        public const int BlankCellsLimit = 0;
        public const ExcelColumn MaxLoadColumn = ExcelColumn.XFD;
        public const int MaxLoadLine = 1048576;

        public Application Application { get; private set; }
        public Workbook Workbook { get; private set; }
        public string FullFilePath { get; private set; }
        public Worksheet CurrentWorksheet { get; private set; }
        public string CurrentSheetName { get; private set; }
        public RangeCoords WorksheetUsedRange { get; private set; }
        public RangeCoords WorksheetFullRange
        {
            get
            {
                var fullRange =  new RangeCoords
                {
                    First = new CellCoords
                    {
                        Row = CurrentWorksheet.Cells[1, 1].Row,
                        Col = (ExcelColumn)CurrentWorksheet.Cells[1, 1].Column,
                    },
                    Last = new CellCoords
                    {
                        Row = CurrentWorksheet.Rows.Count,
                        Col = (ExcelColumn)CurrentWorksheet.Columns.Count,
                    },
                };

                return fullRange;
            }
        }
        public Worksheet[] AllSheets
        {
            get
            {
                var sheets = new List<Worksheet>();
                foreach (Worksheet sheet in Workbook.Sheets)
                {
                    sheets.Add(sheet);
                }
                return sheets.ToArray();
            }
        }
        public string[] AllSheetNames
        {
            get
            {
                return AllSheets.Select(x => x.Name).ToArray();
            }
        }
        public int LastFilledCellRow
        {
            get => LastFilledCell.Row;
        }
        public int LastFilledCellCol
        {
            get => LastFilledCell.Column;
        }
        private Range LastFilledCell
        {
            get
            {
                return CurrentWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            }
        }




        public Excel(string filePath, bool disableMacros = false)
        {
            Application = new Application();
            Application.DisplayAlerts = false; // disable pop-ups when overwriting

            FullFilePath = Ut.GetFullPath(filePath);

            WorksheetUsedRange = new RangeCoords();

            if (disableMacros)
                DisableMacros();

            try
            {
                Workbook = Application.Workbooks.Open(FullFilePath, ReadOnly: (disableMacros && IsMacroSupported()));
            }
            catch (Exception ex)
            {
                throw new ExcelException($"Failed trying open the file. Inner Exception: \n{ex}", ex);
            }
        }

        public Excel(string filePath, string sheetName, bool disableMacros = false)
            : this(filePath, disableMacros)
        {
            CurrentSheetName = sheetName;
            AccessWorksheet(sheetName);
        }




        private bool DisableMacros()
        {
            var disableMacrosOnlyXlsm = IsMacroSupported();

            if (disableMacrosOnlyXlsm) 
                Application.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;

            return disableMacrosOnlyXlsm;
        }

        private void UpdateWorksheetUsedRange()
        {
            // Selects the last cell from Worksheet in first column
            //Range lastRowCell = CurrentWorksheet.Cells[CurrentWorksheet.Rows.Count, 1];
            //var firstRow = lastRowCell.End[XlDirection.xlUp].Row; // press END DOWN_ARROW in sequence to go to the last filled row (or loaded row)
            //var lastRow = lastRowCell.End[XlDirection.xlDown].Row;// press END UP_ARROW in sequence to go to the first filled row

            // Selects the last cell from Worksheet in first row
            //Range lastColCell = CurrentWorksheet.Cells[1, CurrentWorksheet.Columns.Count];
            //var firstCol = (ExcelColumn)lastColCell.End[XlDirection.xlToLeft].Column; // END LEFT_ARROW to go to the first filled cell in the row
            //var lastCol = (ExcelColumn)lastColCell.End[XlDirection.xlToRight].Column; // END RIGHT_ARROW to go to the last filled cell in the row

            //var lastRow = CurrentWorksheet.Rows.Count;
            //var lastCol = (ExcelColumn)CurrentWorksheet.Columns.Count;
            //Range firstCell = CurrentWorksheet.Cells[1, 1];
            //Range lastCell = CurrentWorksheet.Cells[lastRow, (int)lastCol];

            //var firstRow = 0;
            //var firstCol = ExcelColumn.None;
            //foreach (Range cell in CurrentWorksheet.Range[firstCell, lastCell])
            //{
            //    if (!string.IsNullOrWhiteSpace(cell.Value))
            //    {
            //        firstRow = cell.Row;
            //        firstCol = (ExcelColumn)cell.Column;
            //        break;
            //    }
            //}

            //var firstRow = 1;
            //var firstCol = ExcelColumn.A;

            //WorksheetUsedRange = GetUsedRange(GetRange(firstRow, firstCol, lastRow, lastCol), blankCellsLimit: 1000);

            Range firstCellRange = CurrentWorksheet.Cells[1, (int)ExcelColumn.A];

            Range sheetUsedRange = CurrentWorksheet.UsedRange;
            var firstCellUsedRange = sheetUsedRange.Cells[1, (int)ExcelColumn.A];
            var lastCellUsedRange = sheetUsedRange.Cells[sheetUsedRange.Rows.Count, sheetUsedRange.Columns.Count];

            var firstRowFind = CurrentWorksheet.Cells.Find(What: "*",
                                                           After: firstCellRange,
                                                           SearchOrder: XlSearchOrder.xlByRows,
                                                           SearchDirection: XlSearchDirection.xlNext).Row;

            var firstColFind = CurrentWorksheet.Cells.Find(What: "*",
                                                           After: firstCellRange,
                                                           SearchOrder: XlSearchOrder.xlByColumns,
                                                           SearchDirection: XlSearchDirection.xlNext).Column;


            var lastRowFind = CurrentWorksheet.Cells.Find(What: "*", 
                                                          After: firstCellRange,
                                                          SearchOrder: XlSearchOrder.xlByRows,
                                                          SearchDirection: XlSearchDirection.xlPrevious).Row;

            var lastColFind = CurrentWorksheet.Cells.Find(What: "*",
                                                          After: firstCellRange,
                                                          SearchOrder: XlSearchOrder.xlByColumns,
                                                          SearchDirection: XlSearchDirection.xlPrevious).Column;

            var firstRow = (int)firstCellUsedRange.Row < firstRowFind ? (int)firstCellUsedRange.Row : firstRowFind;
            var firstCol = (int)firstCellUsedRange.Column < firstColFind ? (int)firstCellUsedRange.Column : firstColFind;
            var lastRow = (int)lastCellUsedRange.Row < lastRowFind ? (int)lastCellUsedRange.Row : lastRowFind;
            var lastCol = (int)lastCellUsedRange.Column < lastColFind ? (int)lastCellUsedRange.Column : lastColFind;

            WorksheetUsedRange = new RangeCoords(firstRow, (ExcelColumn)firstCol, lastRow, (ExcelColumn)lastCol);
        }

        public void RemoveUnusededFormats()
        {
            Range firstCellRange = CurrentWorksheet.Cells[1, (int)ExcelColumn.A];

            var lastRow = CurrentWorksheet.Cells.Find(What: "*",
                                                      After: firstCellRange,
                                                      SearchOrder: XlSearchOrder.xlByRows,
                                                      SearchDirection: XlSearchDirection.xlPrevious).Row;

            var lastCol = CurrentWorksheet.Cells.Find(What: "*",
                                                      After: firstCellRange,
                                                      SearchOrder: XlSearchOrder.xlByColumns,
                                                      SearchDirection: XlSearchDirection.xlPrevious).Column;

            //WorksheetRange
        }


        // Doesn't work well because to get the real table boundaries it shold process a minimum of rows and columns
        // before stoping at the last loaded one
        private CellCoords[] DetectTableBoundaries(Orientation orientation)
        {
            var detectedCells = new List<CellCoords>();
            Range firstCellRange = CurrentWorksheet.Cells[1, (int)ExcelColumn.A];

            XlDirection xlDirection;
            if (orientation == Orientation.Horizontal)
                xlDirection = XlDirection.xlToRight;
            else if (orientation == Orientation.Vertical)
                xlDirection = XlDirection.xlDown;
            else
                throw new ExcelException($"Orientation '{orientation}' was not mapped for detecting table boundaries.");

            while (true)
            {
                Range currentCellRange = firstCellRange.End[xlDirection];
                var currentCell = new CellCoords(currentCellRange); 

                if (orientation == Orientation.Horizontal && currentCell.Col >= MaxLoadColumn
                    ||
                    orientation == Orientation.Vertical && currentCell.Row >= MaxLoadLine)
                {
                    break;
                }
                else
                {
                    detectedCells.Add(currentCell);
                }
            }

            // if the number of detected boundary cells is odd
            // it's because the first cell (A1) is the first boundary of the first table
            if (!Ut.IsEven(detectedCells.Count))
            {
                detectedCells.Insert(0, new CellCoords(firstCellRange));
            }

            return detectedCells.ToArray();
        }

        public void DetectTables()
        {
            var horizontalBoundaries = DetectTableBoundaries(Orientation.Horizontal);
            var verticalBoundaries = DetectTableBoundaries(Orientation.Vertical);

            // TODO: finish it
        }

        public bool IsMacroSupported()
        {
            return Ut.IsMatch(FullFilePath, @"\.xlsm$");
        }

        private void ReleaseCurrentWorksheet()
        {
            if (CurrentWorksheet != null)
            {
                Marshal.ReleaseComObject(CurrentWorksheet);
            }
        }

        public void AccessWorksheet(string worksheetNameRegex)
        {
            var realSheetName = GetRealWorksheetName(worksheetNameRegex);
            ReleaseCurrentWorksheet();
            try
            {
                //CurrentWorksheet = Workbook.Sheets.Item[CurrentSheetName];
                CurrentWorksheet = Workbook.Worksheets[realSheetName];
                CurrentWorksheet.Activate();
            }
            catch (COMException ex)
            {
                //if (Ut.IsMatch(ex.Message, @"Invalid index\. \(Exception from HRESULT: 0x8002000B \(DISP_E_BADINDEX\)\)"))
                //{
                //    throw new WorksheetNotFoundException(sheetName, filePath);
                //}
                //else
                //{
                //    throw ex;
                //}

                throw new WorksheetNotFoundException(worksheetNameRegex, FullFilePath);
            }

            UpdateWorksheetUsedRange();
        }

        public Range GetRange(string excelRange)
        {
            Range range;
            try
            {
                range = CurrentWorksheet.Range[excelRange];
            }
            catch(COMException ex)
            {
                throw new ExcelException($"The specified range \"{excelRange}\" wasn't accepted by Excel. COMException message: {ex.Message}.", ex);
            }
            
            return range;
        }

        public Range GetRange(string firstCell, string lastCell)
        {
            Range range;
            try
            {
                range = CurrentWorksheet.Range[firstCell, lastCell];
            }
            catch (COMException ex)
            {
                throw new ExcelException($"Either specified range \"{firstCell}\" or \"{lastCell}\" weren't accepted by Excel. COMException message: {ex.Message}.", ex);
            }
            
            return range;
        }

        public Range GetRange(int firstCellRow, ExcelColumn firstCellCol, ExcelColumn lastCellCol)
        {
            return GetRange(firstCellRow, firstCellCol, WorksheetUsedRange.Last.Row, lastCellCol);
        }

        public Range GetRange(int firstCellRow, ExcelColumn firstCellCol, int lastCellRow, ExcelColumn lastCellCol)
        {
            return GetRange(firstCellRow, (int)firstCellCol, lastCellRow, (int)lastCellCol);
        }

        public Range GetRange(int firstCellRow, int firstCellCol, int lastCellRow, int lastCellCol)
        {
            Range firstCell;
            Range lastCell;
            Range range;

            if (firstCellRow <= 0 | firstCellCol <= 0 | lastCellRow <= 0 | lastCellCol <= 0)
            {
                throw new RpaLibArgumentException($"Row and Column excel indexes start must be positive integers starting from 1.");
            }

            // parse the first cell
            try
            {
                firstCell = CurrentWorksheet.Cells[firstCellRow, firstCellCol];
            }
            catch (COMException ex)
            {
                throw new ExcelException($"The specified row [{firstCellRow}] or column [{firstCellCol}] for the first cell weren't accepted by Excel. COMException message: {ex.Message}", ex);
            }

            // parse the last cell
            try
            {
                lastCell = CurrentWorksheet.Cells[lastCellRow, lastCellCol];
            }
            catch (COMException ex)
            {
                throw new ExcelException($"The specified row [{lastCellRow}] or column [{lastCellCol}] for the last cell weren't accepted by Excel. COMException message: {ex.Message}", ex);
            }

            // get full range between first and last cell in left-right up-down order in a reading like manner.

            try
            {
                range = CurrentWorksheet.Range[firstCell, lastCell];
            }
            catch (COMException ex)
            {
                throw new ExcelException($"The specified range between cells ({firstCellRow}, {firstCellCol}) and ({lastCellRow}, {lastCellCol}) wasn't accepted by Excel. COMException message: {ex.Message}");
            }

            return range;
        }

        public RangeCoords GetUsedRange(string excelRange)
        {
            return GetUsedRange(GetRange(excelRange));
        }

        private RangeCoords GetUsedRange(Range range, int blankCellsLimit = BlankCellsLimit)
        {
            // doesn't work: gets the last column of the sheet instead of the last column from range
            //var lastCell = range.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

            int firstRow;
            int firstCol;
            Range firstCell;

            try
            {
                firstCell = range.Cells[1];
            }
            catch
            {
                try
                {
                    firstCell = range.Cells[1, 1];
                }
                catch (Exception ex)
                {
                    throw new ExcelException($"Could not discover which is the first cell of the Range. Original exception:\n{ex}");
                }
            }
                

            try
            {
                firstRow = firstCell.Row;
            }
            catch (Exception ex)
            {
                throw new ExcelException($"Could not discover which is the first cell's row of the Range. Inner Exception:\n{ex}", ex);
            }

            try
            {
                firstCol = firstCell.Column;
            }
            catch (Exception ex)
            {
                throw new ExcelException($"Could not discover which is the first cell's column of the Range. Inner Exception:\n{ex}", ex);
            }

            var biggestRow = 0;
            var biggestCol = 0;

            foreach (Range row in range.Rows)
            {
                StringBuilder rowValue = new StringBuilder();

                var blankCellsCount = 0;
                foreach (Range cell in row.Cells)
                {
                    var cellValue = cell.Value?.ToString() ?? string.Empty;

                    // stop at first blank cell of row
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        rowValue.Append(cellValue);
                    }
                    else
                    {
                        blankCellsCount++;
                        if (blankCellsCount > blankCellsLimit)
                            break;
                    }
                }

                var fullRow = rowValue.ToString();

                if (!string.IsNullOrWhiteSpace(fullRow) && row.Row > biggestRow)
                    biggestRow = row.Row;

                // stop if the current full row is all blank
                if (string.IsNullOrWhiteSpace(fullRow))
                    break;
            }

            foreach (Range col in range.Columns)
            {
                var colValue = new StringBuilder();

                var blankCellsCount = 0;
                foreach (Range cell in col.Cells)
                {
                    var cellValue = cell.Value?.ToString() ?? string.Empty;

                    // stop at first blank cell of column
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        colValue.Append(cellValue);
                    }
                    else
                    {
                        blankCellsCount++;
                        if (blankCellsCount > blankCellsLimit)
                            break;
                    }
                }
                    
                var fullColumn = colValue.ToString();

                if (!string.IsNullOrWhiteSpace(fullColumn) && col.Column > biggestCol) 
                    biggestCol = col.Column;

                // stop if the current full column is all blank
                if (string.IsNullOrWhiteSpace(fullColumn))
                    break;
            }

            var usedRange = new RangeCoords
            {
                First = new CellCoords
                {
                    Row = firstRow,
                    Col = (ExcelColumn)firstCol,
                },
                Last = new CellCoords
                {
                    Row = biggestRow,
                    Col = (ExcelColumn)biggestCol,
                },
            };

            return usedRange;
        }

        public bool IsRangeSingleCell(Range range)
        {
            return range.Count == 1;
        }

        public void CustomFormatRange(string excelRange, string customFormat)
        {
            CustomFormatRange(GetRange(excelRange), customFormat);
        }

        public void CustomFormatRange(string firstCell, string lastCell, string customFormat)
        {
            CustomFormatRange(GetRange(firstCell, lastCell), customFormat);
        }

        public void CustomFormatRange(Range range, string customFormat)
        {
            try
            {
                if (IsRangeSingleCell(range))
                {
                    range.NumberFormat = customFormat;
                }
                else
                {
                    foreach (Range cell in range)
                        cell.NumberFormat = customFormat;
                }
            }
            catch (COMException ex)
            {
                throw new ExcelException("Could not set custom format in range", ex);
            }
            
        }

        public void DeleteCells(string excelRange)
        {
            DeleteCells(GetRange(excelRange));
        }

        private void DeleteCells(Range range)
        {
            try
            {
                range.Delete(XlDeleteShiftDirection.xlShiftUp);
            }
            catch(Exception ex)
            {
                throw new ExcelException("Error when trying to delete all cells from a Range object of Excel interop." +
                    $" Inner Exception: {ex}", ex);
            }
        }

        // usage: InsertFormula("=VLOOKUP()", "A1:B20");
        public void InsertFormula(string formula, string excelRange)
        {
            Range formulaFullRange = GetRange(excelRange);
            InsertFormula(formulaFullRange, formula);
        }

        // usage: InsertFormula("=VLOOKUP()", "A1", "B20");
        public void InsertFormula(string formula, string firstCell, string lastCell)
        {
            Range formulaFullRange = GetRange(firstCell, lastCell);
            InsertFormula(formulaFullRange, formula);
        }

        // To work with sheetName used ranges.
        // usage:
        //    InsertFormula("=VLOOKUP()", 1, 1, 20, 2); // insert formula in range "A1:B20"
        //    InsertFormula("=VLOOKUP()", 1, 1); //last cell will be the last used firstCellRow and column
        //    InsertFormula("=VLOOKUP()", 1, 1, 20); // last cell column will be the last used
        //    InsertFormula("=VLOOKUP()", 1, 1, lastCellCol: 2); // last cell firstCellRow will be the last used
        public void InsertFormula(string formula, int firstCellRow, ExcelColumn firstCellCol, int lastCellRow = 0, ExcelColumn lastCellCol = ExcelColumn.None)
        {
            if (lastCellRow <= 0)
                lastCellRow = WorksheetUsedRange.Last.Row;

            if (lastCellCol <= 0)
                lastCellCol = WorksheetUsedRange.Last.Col;

            Range formulaFullRange = GetRange(firstCellRow, firstCellCol, lastCellRow, lastCellCol);
            InsertFormula(formulaFullRange, formula);
        }

        public void InsertFormula(string formula, int firstCellRow, int firstCellCol, int lastCellRow = 0, int lastCellCol = 0)
        {
            InsertFormula(formula, firstCellRow, (ExcelColumn)firstCellCol, lastCellRow, (ExcelColumn)lastCellCol);
        }

        private void InsertFormula(Range range, string excelFormula)
        {
            try
            {
                range.Formula = excelFormula;
            }
            catch (COMException ex)
            {
                throw new ExcelException("Excel: Error trying to insert formula in cell. Check formula syntax.", ex);
            }
            
        }

        public string ExtractFormula(string cell)
        {
            Range cellRange = GetRange(cell);
            
            string extractedFormula = (string)cellRange.Formula;

            return extractedFormula;
        }

        public void RemoveFormulaKeepValue(string cellRange)
        {
            Range range = GetRange(cellRange);

            foreach (Range cell in range)
            {
                if (cell.HasFormula)
                    cell.Value = cell.Value;
            }
        }

        public void Quit()
        {
            try
            {
                if (Application != null)
                    Application.DisplayAlerts = false; // disable confirmation popups

                //Workbook.Close(SaveChanges: false);
                ReleaseCurrentWorksheet();

                if (Workbook != null)
                    Marshal.ReleaseComObject(Workbook);
                
                
                if (Application != null)
                {
                    Application.Quit(); // Maybe an interop bug: just puts in background instead of closing.
                    Marshal.ReleaseComObject(Application);
                }

                CurrentWorksheet = null;
                Workbook = null;
                Application = null;

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
                //Ut.KillProcess(ProcessName);
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

        //public void SetActiveSheet(dynamic nameOrIndex)
        //{
        //    if (nameOrIndex is string || nameOrIndex is int)
        //        _ = Workbook.Sheets.Item[nameOrIndex];
        //    else
        //        throw new RpaLibArgumentException("SetActiveSheet() receives nameOrIndex of types int or string only.");
        //}

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

        //public string ReadCellTest(int row, int col)
        //{
        //    return (string)Application.Cells[row, col].Text;
        //}

        public DataTable ReadAll()
        {
            return ReadAll(WorksheetUsedRange.First.Row, WorksheetUsedRange.First.Col, WorksheetUsedRange.Last.Row, WorksheetUsedRange.Last.Col);
        }

        public DataTable ReadAll(int startRow, ExcelColumn startCol)
        {
            return ReadAll(startRow, startCol, WorksheetUsedRange.Last.Row, WorksheetUsedRange.Last.Col);
        }

        public DataTable ReadAll(int startRow, ExcelColumn startCol, ExcelColumn endCol)
        {
            return ReadAll(startRow, startCol, WorksheetUsedRange.Last.Row, endCol);
        }

        public DataTable ReadAll(ExcelColumn startCol, ExcelColumn endCol)
        {
            return ReadAll(WorksheetUsedRange.First.Row, startCol, WorksheetUsedRange.Last.Row, endCol);
        }

        public DataTable ReadAll(SheetInfo sheet, bool breakAtEmptyLine = true)
        {
            return ReadAll(sheet.FirstRow, sheet.FirstCol, sheet.LastRow, sheet.LastCol, breakAtEmptyLine);
        }

        public DataTable ReadAll(int startRow, ExcelColumn startCol, int endRow, ExcelColumn endCol, bool breakAtEmptyLine = true)
        {
            var usedRange = GetUsedRange(GetRange(startRow, startCol, endRow, endCol));

            var contents = ReadCell(usedRange.First.Row, (int)usedRange.First.Col,
                                          usedRange.Last.Row, (int)usedRange.Last.Col,
                                          breakAtEmptyLine: breakAtEmptyLine);

            return ArrayStrTableToDataTable(contents);
        }

        public Dictionary<string, DataTable> ReadAllSheets(SheetInfo[] sheets, bool breakAtEmptyLine = true)
        {
            var sheetsRead = new Dictionary<string, DataTable>();

            for (var i = 0; i < AllSheetNames.Length; i++)
            {
                var sheetInfo = sheets[i];

                var currentSheetName = sheetInfo.Name == DefaultSheetName ? AllSheetNames[i] : sheetInfo.Name;
                AccessWorksheet(currentSheetName);

                // if SheetInfo don't define sheet last row and column use the last used in sheet
                var endRow = sheetInfo.LastRow <= 0 ? WorksheetUsedRange.Last.Row : sheetInfo.LastRow;
                var endCol = sheetInfo.LastCol <= ExcelColumn.None ? WorksheetUsedRange.Last.Col : sheetInfo.LastCol;

                // get the actual used range based on sheetInfo parameters
                var usedRange = GetUsedRange(GetRange(sheetInfo.FirstRow, sheetInfo.FirstCol, endRow, endCol));


                var contents = ReadCell(sheetInfo.FirstRow, (int)sheetInfo.FirstCol,
                                                usedRange.Last.Row, (int)usedRange.Last.Col,
                                                breakAtEmptyLine: breakAtEmptyLine);

                var dtContents = ArrayStrTableToDataTable(contents);

                sheetsRead.Add(currentSheetName, dtContents);
            }

            return sheetsRead;
        }

        public string ReadCell(int row, int col)
        {
            return (string)CurrentWorksheet.Rows.Item[row].Columns.Item[col].Text;
        }

        public string[] ReadCell(int row, int[] cols)
        {
            List<string> cells = new List<string>();

            foreach (int col in cols)
                cells.Add((string)CurrentWorksheet.Rows.Item[row].Columns.Item[col].Text);

            return cells.ToArray();
        }

        public string[] ReadCell(int[] rows, int col)
        {
            List<string> cells = new List<string>();

            foreach (int row in rows)
                cells.Add((string)CurrentWorksheet.Rows.Item[row].Columns.Item[col].Text);

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

                var rowRange = Ut.Seq(minRow, maxRow);

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
                            line.Add((string)CurrentWorksheet.Rows.Item[row].Columns.Item[col].Text);
                            break;
                        }
                        catch (COMException ex)
                        {
                            if (Ut.IsMatch(ex.Message, @"The message filter indicated that the application is busy\."))
                                continue;
                            else
                                throw new ExcelException($"Could not read the cell value. Cell[{row}, {col}]", ex);
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
            return ReadCell(Ut.Seq(startRow, endRow), Ut.Seq(startCol, endCol), breakAtEmptyLine);
        }

        public string[] ReadCell(int row, int startCol, int endCol)
        {
            return ReadCell(row, Ut.Seq(startCol, endCol));
        }

        private string GetRealWorksheetName(string sheetNameRegex)
        {
            var realSheetName = AllSheetNames.Where(x => Ut.IsMatch(x, sheetNameRegex)).FirstOrDefault();
            if (realSheetName == null)
                throw new ExcelException($"The worksheet named \"{sheetNameRegex}\" was not found.");

            return realSheetName;
        }

        public void WriteNextFreeRow(DataTable table)
        {
            WriteNextFreeRow(table, WorksheetUsedRange.First.Col);
        }

        public void WriteNextFreeRow(DataTable table, ExcelColumn startCol)
        {
            int startRow = WorksheetUsedRange.Last.Row + 1;

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
            WriteCell(startRow, startCol, rowList.ToArray());
        }

        public void WriteCell(int row, ExcelColumn col, object value)
        {
            WriteCell(row, (int)col, value);
        }

        public void WriteCell(int row, int col, object value)
        {
            CurrentWorksheet.Rows.Item[row].Columns.Item[col] = value;
            UpdateWorksheetUsedRange();
        }

        public void WriteCell(int firstCellRow, ExcelColumn firstCellCol, object[] values, InsertMethod rowOrCol)
        {
            WriteCell(firstCellRow, (int)firstCellCol, values, rowOrCol);
        }

        public void WriteCell(int firstCellRow, int firstCellCol, object[] values, InsertMethod rowOrCol)
        {
            for (int j = 0; j < values.Count(); j++)
            {
                switch (rowOrCol)
                {
                    case InsertMethod.AsRow:
                        CurrentWorksheet.Rows.Item[firstCellRow].Columns[firstCellCol + j] = values[j];
                        break;
                    case InsertMethod.AsColumn:
                        CurrentWorksheet.Rows.Item[firstCellRow + j].Columns[firstCellCol] = values[j];
                        break;
                    default:
                        throw new RpaLibArgumentException($"Invalid insertion method: {rowOrCol}");
                }
            }
            UpdateWorksheetUsedRange();
        }

        public void WriteCell(int firstCellRow, ExcelColumn firstCellCol, object[][] tableOfValues)
        {
            WriteCell(firstCellRow, (int)firstCellCol, tableOfValues);
        }

        public void WriteCell(int firstCellRow, int firstCellCol, object[][] tableOfValues)
        {
            for (int i = 0; i < tableOfValues.Count(); i++)
                WriteCell(i + firstCellRow, firstCellCol, tableOfValues[i], InsertMethod.AsRow);
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
                startRow = excel.WorksheetUsedRange.Last.Row + 1;

            if (appendCol)
                startCol = (int)excel.WorksheetUsedRange.Last.Col + 1;

            if (endRow < 1)
                endRow = excel.WorksheetUsedRange.Last.Row;

            if (endCol < 1)
                endCol = (int)excel.WorksheetUsedRange.Last.Col;

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
                        excel.WorksheetUsedRange.Last.Row,
                        (int)excel.WorksheetUsedRange.Last.Col);
            else if (readRow > 0)
                returnValue =
                    excel.ReadCell(readRow, Ut.Seq(startCol, endCol));
            else if (readCol > 0)
                returnValue =
                    excel.ReadCell(Ut.Seq(startRow, endRow), readCol);

            excel.Save(saveAs);
            excel.Quit();

            return returnValue;
        }

        public static DataTable ReadAll(string filePath, SheetInfo sheetInfo, bool visible = false, bool disableMacros = false, bool breakAtEmptyLine = true)
        {
            Excel excel = new Excel(filePath, disableMacros);
            excel.ToggleVisible(visible);
            excel.AccessWorksheet(sheetInfo.Name);

            var sheetInfoFixed = FixReadAllInput(sheetInfo, excel);
            var content = excel.ReadAll(sheetInfoFixed, breakAtEmptyLine);

            excel.Quit();

            return content;
        }

        public static DataTable ReadAll(string filePath, string sheetNameRegex, bool visible = false, bool disableMacros = false, bool breakAtEmptyLine = true,
            int startRow = 1, ExcelColumn startCol = ExcelColumn.A, int endRow = 0, ExcelColumn endCol = ExcelColumn.None)
        {
            Excel excel = new Excel(filePath, disableMacros);
            excel.ToggleVisible(visible);
            excel.AccessWorksheet(sheetNameRegex);

            var sheetInfo = FixReadAllInput(startRow, startCol, endRow, endCol, excel);
            var content = excel.ReadAll(sheetInfo, breakAtEmptyLine);

            excel.Quit();

            return content;
        }

        private static SheetInfo FixReadAllInput(SheetInfo sheetInfo, Excel excel)
        {
            return FixReadAllInput(sheetInfo.FirstRow, sheetInfo.FirstCol, sheetInfo.LastRow, sheetInfo.LastCol, excel);
        }

        private static SheetInfo FixReadAllInput(int startRow, ExcelColumn startCol, int endRow, ExcelColumn endCol, Excel excel)
        {
            var sheetInfo = new SheetInfo();

            if (startRow < 1)
                sheetInfo.FirstRow = 1;
            else
                sheetInfo.FirstRow = startRow;

            if (startCol < ExcelColumn.A)
                sheetInfo.FirstCol = ExcelColumn.A;
            else
                sheetInfo.FirstCol = startCol;

            if (endRow <= 0)
                sheetInfo.LastRow = excel.WorksheetFullRange.Last.Row;
            else
                sheetInfo.LastRow = endRow;

            if (endCol <= ExcelColumn.None)
                sheetInfo.LastCol = excel.WorksheetFullRange.Last.Col;
            else
                sheetInfo.LastCol = endCol;

            return sheetInfo;
        }

        public static Dictionary<string, DataTable> ReadAllSheets(string filePath, bool visible = false, bool disableMacros = false, bool breakAtEmptyLine = true)
        {
            var excel = new Excel(filePath, disableMacros);
            var sheetsInfo = excel.AllSheetNames.Select(x => new SheetInfo() { Name = x }).ToArray();
            return ReadAllSheets(excel, sheetsInfo, visible, breakAtEmptyLine);
        }

        public static Dictionary<string, DataTable> ReadAllSheets(string filePath, SheetInfo[] sheets, bool visible = false, bool disableMacros = false, bool breakAtEmptyLine = true)
        {
            var excel = new Excel(filePath, disableMacros);
            return ReadAllSheets(excel, sheets, visible, breakAtEmptyLine);
        }

        private static Dictionary<string, DataTable> ReadAllSheets(Excel excel, SheetInfo[] sheets, bool visible = false, bool breakAtEmptyLine = true)
        {
            excel.ToggleVisible(visible);

            var contentsAllSheets = excel.ReadAllSheets(sheets, breakAtEmptyLine);

            excel.Quit();

            return contentsAllSheets;
        }

        // Writes an array of values as a column or firstCellRow into the sheetName. It starts at the cell specified by firstCellRow and firstCellCol values.
        // If firstCellRow or firstCellCol are negative, it will consider the next free one. 
        public static void WriteAll(string filePath, string sheetName, int row, int col, string[] values, InsertMethod rowOrCol, bool visible = false)
        {
            var excel = new Excel(filePath, sheetName, disableMacros: false);
            excel.ToggleVisible(visible);

            if (row < 0)
                row = excel.WorksheetUsedRange.Last.Row;

            if (col < 0)
                col = (int)excel.WorksheetUsedRange.Last.Col;

            excel.WriteCell(row, col, values, rowOrCol);
            excel.Save();
            excel.Quit();
        }

        // TODO: code validation for excel Ut.Seq (accept only: A1, A1:B2, A:A)
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

        public static void WriteNextFreeRow(string filePath, DataTable table, ExcelColumn startCol = ExcelColumn.None, string sheetName = DefaultSheetName, bool visible = false, bool disableMacros = false)
        {
            string fileFullPath = Ut.GetFullPath(filePath);
            Excel excel = new Excel(fileFullPath, sheetName, disableMacros);

            excel.ToggleVisible(visible);

            if (startCol == ExcelColumn.None)
                excel.WriteNextFreeRow(table);
            else
                excel.WriteNextFreeRow(table, startCol);

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
            if (arrayStrTable.Count() == 0)
            {
                throw new RpaLibException("Array string table came with 0 rows.");
            }
            else if (arrayStrTable == null)
            {
                throw new RpaLibException("Array string table came nulled.");
            }
            else if (arrayStrTable.Where(x => x?.Count() == 0).Count() > 1) {
                throw new RpaLibException("Array string table came with 0 columns in, at least, one row.");
            }
            else if (arrayStrTable.Where(x => x == null).Count() > 1)
            {
                throw new RpaLibException("Array string table came with null in, at least, one row.");
            }

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