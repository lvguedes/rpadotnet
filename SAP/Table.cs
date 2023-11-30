using RpaLib.ProcessAutomation;
using sapfewse;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Tracing;

namespace RpaLib.SAP
{

    public class Table : SapComponent<GuiTableControl>
    {
        const int DELAY_AFTER_SCROLL = 0;
        public SapComWrapper<GuiTableControl> GuiTableControl
        {
            get => Session.FindById<GuiTableControl>(PathId, suppressTrace: true);
        }

        private DataTable _dt;
        public DataTable DataTable
        {
            get
            {
                //Parse();
                return _dt;
            }
            private set
            {
                _dt = value;
            }
        }

        public int CurrentRow { get; private set; }

        public int RowCount { get; private set; }
        
        public int VisibleRowCount { get; private set; }

        public int DelayAfterScroll { get; }

        public Scroll VerticalScrollbar
        {
            get => new Scroll(GuiTableControl.Com.VerticalScrollbar, DelayAfterScroll);
        }

        public Scroll HorizontalScrollbar
        {
            get => new Scroll(GuiTableControl.Com.HorizontalScrollbar, DelayAfterScroll);
        }

        public Table(Session session, string fullPathId, string selectRegex = null, int regexMatches = 1, int[] selectRows = null, int delayAfterScroll = DELAY_AFTER_SCROLL)
            : this(session, name: string.Empty, fullPathId, selectRegex, regexMatches, selectRows, delayAfterScroll)
        { }

        public Table(Session session, string name, string fullPathId, string selectRegex = null, int regexMatches = 1, int[] selectRows = null, int delayAfterScroll = DELAY_AFTER_SCROLL)
            : base(session, fullPathId)
        {
            Name = name;
            DelayAfterScroll = delayAfterScroll;
            DataTable = new DataTable();
            //RefreshTableObj();
            Parse(selectRegex, regexMatches, selectRows);
            Trace.WriteLine(Info());
        }

        /* Remove in future
        private void RefreshTableObj()
        {
            Trace.WriteLine("Refreshing the GuiTableControl object...");
            //GuiTableControl = Session.FindById<GuiTableControl>(FullPathId);
            CurrentRow = GuiTableControl.CurrentRow;
            RowCount = GuiTableControl.RowCount;
            VisibleRowCount = GuiTableControl.VisibleRowCount;
        }
        */

        
        /* Remove in furure
        private void ResetScrolling()
        {
            Trace.WriteLine("Resetting scrollings' position to 0...");
            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.VerticalScrollbar.Position = 0;
                },
                () => { }
            );

            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.HorizontalScrollbar.Position = 0;
                },
                () => { }
            );
        }
        */

        public string Info() => Info(this);
        public static string Info(Table table)
        {
            //table.RefreshTableObj();
            return
                string.Join(Environment.NewLine,
                    $"Table \"{table.Name}\" captured:",
                    $"  CharHeight: \"{table.GuiTableControl.Com.Text}\"",
                    $"  CharWidth:  \"{table.GuiTableControl.Com.Text}\"",
                    $"  CharTop:    \"{table.GuiTableControl.Com.Text}\"",
                    $"  CurrentCol: \"{table.GuiTableControl.Com.Text}\"",
                    $"  CurrentRow: \"{table.GuiTableControl.Com.Text}\"",
                    $"  RowCount:   \"{table.GuiTableControl.Com.Text}\"",
                    $"  Rows.Count: \"{table.GuiTableControl.Com.Rows.Count}\"",
                    $"  VisibleRowCount: \"{table.GuiTableControl.Com.Text}\"",
                    $"  HorizontalScrollbar:",
                    $"    Minimum:  \"{table.GuiTableControl.Com.HorizontalScrollbar.Minimum}\"",
                    $"    Maximum:  \"{table.GuiTableControl.Com.HorizontalScrollbar.Maximum}\"",
                    $"    Position: \"{table.GuiTableControl.Com.HorizontalScrollbar.Position}\"",
                    $"    PageSize: \"{table.GuiTableControl.Com.HorizontalScrollbar.PageSize}\"",
                    $"  VerticalScrollBar:",
                    $"    Minimum:  \"{table.GuiTableControl.Com.VerticalScrollbar.Minimum}\"",
                    $"    Maximum:  \"{table.GuiTableControl.Com.VerticalScrollbar.Maximum}\"",
                    $"    Position: \"{table.GuiTableControl.Com.VerticalScrollbar.Position}\"",
                    $"    PageSize: \"{table.GuiTableControl.Com.VerticalScrollbar.PageSize}\"",
                    $" --- Table object methods ---"
                    , $"Filled Rows: {table._dt.Rows.Count}"
                    //,$"  FulfilledRowsCount(): {table.FulfilledRowsCount()}"
                    //,$"  IsEmpty(): {table.IsEmpty()}"
                    );
        }

        public static string RowInfo(GuiTableRow row)
        {
            return string.Join(Environment.NewLine,
                $"The first row from table above was captured:)",
                $"  Count:      \"{row.Count}\" (number of cells in the row)",
                $"  Type:       \"{row.Type}\"",
                $"  Selectable: \"{row.Selectable}\" (if the row can be selected)",
                $"  Selected:   \"{row.Selected}\" (if the row is selected)");
        }

        /* Remove in future
        private void ScrollDown(int row)
        {
            int oldPos = GuiTableControl.VerticalScrollbar.Position;
            // relaxed scrolling
            TryEmptyCellActionRelax(
                () =>
                {
                    //rpa.MessageBox("Checking if scroll is needed");
                    Trace.WriteLine($"Checking if scroll down is needed for row {row}");
                    if (row % GuiTableControl.VisibleRowCount == 0 && row != 0)
                    {
                        //rpa.MessageBox("Will scroll down now");
                        Trace.WriteLine("Will scroll down now...");

                        GuiTableControl.VerticalScrollbar.Position += GuiTableControl.VisibleRowCount;
                        //RefreshTableObj();

                        Trace.WriteLine(string.Join(Environment.NewLine,
                            $"Scrolling down...",
                            $"  Old position: {oldPos}",
                            $"  New position: {GuiTableControl.VerticalScrollbar.Position}"));
                        //$"  New position: {newPos}"));
                    }
                });

            //RefreshTableObj(); // for the case it cannot change position
        }
        */

        private void TryEmptyCellActionRelax(Action tryBlock)
        {
            TryEmptyCellActionRelax(tryBlock, () => { });
        }
        private void TryEmptyCellActionRelax(Action tryBlock, Action catchBlock)
        {
            try
            {
                tryBlock();
            }
            catch (COMException ex)
            {
                // Add more messages here if different new messages appear for the same error when trying to access a Text property when 
                // element is empty or there is no element at all.
                if (ex.Message.Equals("The server threw an exception. (Exception from HRESULT: 0x80010105 (RPC_E_SERVERFAULT))")
                    || ex.Message.Equals("The method got an invalid argument."))
                {
                    //RefreshTableObj();
                    catchBlock();
                }
                else
                {
                    Trace.WriteLine("Unknown COMException occurred." +
                        " It's different from the exception that signals empty cells." +
                        $" See:\n{ex}");
                    throw ex;
                }
            }
        }

        private Dictionary<string, int> GetTableCounters()
        {
            Dictionary<string, int> counters = new Dictionary<string, int>();
            // get columns count or zero if table is empty

            //RefreshTableObj();
            TryEmptyCellActionRelax(
                () =>
                {
                    counters["columns"] = GuiTableControl.Com.Columns.Count;
                },
                () =>
                {
                    counters["columns"] = 0;
                });

            TryEmptyCellActionRelax(
                () =>
                {
                    counters["rows"] = GuiTableControl.Com.RowCount;
                },
                () =>
                {
                    counters["rows"] = 0;
                });

            return counters;
        }

        public GuiVComponent GetCell(int row, int column)
        {
            // The native index is always relative to the number of visible rows 
            // in the screen
            // This function makes it possible to use the real index from any collection
            // to get the cell
            return GuiTableControl.Com.GetCell(row % GuiTableControl.Com.VisibleRowCount, column);
        }

        public void MakeRowVisible(int row)
        {
            int currentPage = GuiTableControl.Com.VerticalScrollbar.Position / GuiTableControl.Com.VisibleRowCount;

            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.Com.VerticalScrollbar.Position = row;
                });
            //RefreshTableObj();
        }

        public GuiTableRow GetRow(int rowIndex)
        {
            // put the rowIndex at the top of table after scrolling
            MakeRowVisible(rowIndex);

            //Trace.WriteLine($"Will get row {rowIndex} now...");
            //return GuiTableControl.Rows.ElementAt(rowIndex % GuiTableControl.VisibleRowCount);

            // Get the top row of table
            return GuiTableControl.Com.Rows.ElementAt(0) as GuiTableRow;
        }

        public GuiTableRow GetRow(string cellValueRegex)
        {
            var foundRowIndex = _dt.Rows.Cast<DataRow>()
                .Select((row, index) => (row, index))
                .Where(
                    x => x.row.ItemArray
                        .Where(y => Ut.IsMatch((string)y, cellValueRegex)).FirstOrDefault() != null)
                .FirstOrDefault().index;

            return GetRow(foundRowIndex);
        }

        public void SelectRow(int rowIndex) => GetRow(rowIndex).Selected = true;

        public void SelectRow(string cellValueRegex) => GetRow(cellValueRegex).Selected = true;

        public bool IsRowSelected(int rowIndex) => GetRow(rowIndex).Selected;

        public bool IsRowSelected(string cellValueRegex) => GetRow(cellValueRegex).Selected;

        public bool IsRowSelectable(int rowIndex) => GetRow(rowIndex).Selectable;

        public bool IsRowSelectable(string cellValueRegex) => GetRow(cellValueRegex).Selectable;

        public void SelectAllCols() => GuiTableControl.Com.SelectAllColumns();

        public void DeselectAllCols() => GuiTableControl.Com.DeselectAllColumns();

        /* Remove in future
        public int FulfilledRowsCount()
        {
            int fulfilledRows = 0;
            //Dictionary<string, int> counters = GetTableCounters();

            for (int row = 0; row < GetTableCounters()["rows"]; row++)
            {
                //rpa.MessageBox($"Processing line {row}");
                ScrollDown(row);

                for (int col = 0; col < GetTableCounters()["columns"]; col++)
                {
                    Trace.WriteLine($"Iteration: row {row}, col {col}");
                    bool isCellNullOrEmpty = false;
                    TryEmptyCellActionRelax(
                        () =>
                        {
                            if (!string.IsNullOrEmpty(
                                     GetCell(row, col).Text))
                            {
                                fulfilledRows++;
                                isCellNullOrEmpty = true;
                            }
                        },
                        () =>
                        {
                            isCellNullOrEmpty = true;
                        });
                    if (isCellNullOrEmpty) break;
                    //
                    //try
                    //{
                    //    if (!string.IsNullOrEmpty(GuiTableControl.GetCell(row, col).Text))
                    //    {
                    //        fulfilledRows++;
                    //        break;
                    //    }
                    //}
                    //catch(COMException ex)
                    //{
                    //    if (ex.Message.Equals("The server threw an exception. (Exception from HRESULT: 0x80010105 (RPC_E_SERVERFAULT))"))
                    //    {
                    //        Trace.WriteLine("Caught a null empty cell. Skipping...");
                    //        RefreshTableObj();
                    //        break;
                    //    }
                    //    else
                    //    {
                    //        Trace.WriteLine("Unknown COMException occurred." +
                    //            " It's different from the exception that signals empty cells." +
                    //            $" See:\n{ex}");
                    //        throw ex;
                    //    }
                    //}
                    //
                }

                // if last line is empty, stop analyzing lines
                if (fulfilledRows != row + 1)
                    break;
            }

            ResetScrolling();

            return fulfilledRows;
        }
*/

        /* Remove in future
        public bool IsEmpty()
        {
            bool isEmpty = false;

            if (FulfilledRowsCount() == 0)
                isEmpty = true;

            Trace.WriteLine($"Table \"{Name}\" is empty? {isEmpty}");

            return isEmpty;
        }
        */

        /* Remove in future
        public void ParseOld()
        {
            RefreshTableObj();
            _dt = new DataTable();

            foreach (GuiTableColumn col in GuiTableControl.Columns)
            {
                _dt.Columns.Add(col.Title, typeof(string));
            }

            int fulfilledRowsCount = FulfilledRowsCount();
            for (int row = 0; row < fulfilledRowsCount; row++)
            {
                ScrollDown(row);

                DataRow datarow = _dt.NewRow();
                for (int col = 0; col < GetTableCounters()["columns"]; col++)
                {
                    //rpa.MessageBox($"Getting cell ({row}, {col})...");
                    datarow[col] = GetCell(row, col).Text;
                }
                _dt.Rows.Add(datarow);
            }
            ResetScrolling();
        }
        */

        public void Parse(string selectRegex = null, int regexMatches = 1, int[] selectRows = null)
        {
            _dt = new DataTable();

            foreach (GuiTableColumn col in GuiTableControl.Com.Columns)
            {
                _dt.Columns.Add(col.Title, typeof(string));
            }

            int fulfilledRowsCount = VerticalScrollbar.Maximum + 1;
            for (int row = 0; row < fulfilledRowsCount; row++)
            {
                //if (row != 0 && row % VerticalScrollbar.PageSize == 0)
                if (VerticalScrollbar.IsNeeded(row))
                    VerticalScrollbar.NextPage();

                DataRow datarow = _dt.NewRow();
                for (int col = 0; col < GetTableCounters()["columns"]; col++)
                {
                    //rpa.MessageBox($"Getting cell ({row}, {col})...");
                    var currentCell = GetCell(row, col);
                    datarow[col] = currentCell.Text;

                    if (selectRegex != null && Ut.IsMatch(currentCell.Text, selectRegex))
                    {
                        _dt.Rows.Add(datarow);
                        //currentCell.SetFocus();
                        SelectRow(row);
                        regexMatches--;
                        if (regexMatches == 0)
                            return;
                        else
                            continue;
                    }
                    else if (selectRows != null && selectRows.Contains(row) && selectRows.Length > 1)
                    {
                        if (selectRows.Length > 1)
                        {
                            SelectRow(row);
                            var selectQuickRowsList = selectRows.ToList();
                            selectQuickRowsList.RemoveAt(row);
                            selectRows = selectQuickRowsList.ToArray();
                            continue;
                        }
                        else
                        {
                            _dt.Rows.Add(datarow);
                            SelectRow(row);
                            return;
                        }
                    }
                }
                _dt.Rows.Add(datarow);
            }
            VerticalScrollbar.Reset();
        }

        public void PrintDataTable()
        {
            Trace.WriteLine($"Printing the DataTable that represents SAP table \"{Name}\":", color: ConsoleColor.Yellow);
            Trace.WriteLine(Ut.PrintDataTable(_dt), withTimeSpec: false, color: ConsoleColor.Magenta);
        }
    }
}
