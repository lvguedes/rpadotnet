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

    public class Table : SapComponent, ISapTabular
    {
        //public Session Session { get; private set; }
        public GuiTableControl GuiTableControl { get; private set; }

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

        public Table(Session session)
            : this(session: session, name: string.Empty, fullPathId: string.Empty)
        { }

        public Table(Session session, string fullPathId)
            : this(session, name: string.Empty, fullPathId)
        { }

        public Table(Session session, string name, string fullPathId) : base(null)
        {
            Session = session;
            Name = name;
            FullPathId = fullPathId;
            DataTable = new DataTable();
            RefreshTableObj();
            Parse();
            Trace.WriteLine(Info());
        }

        private void RefreshTableObj()
        {
            Trace.WriteLine("Refreshing the GuiTableControl object...");
            GuiTableControl = Session.FindById<GuiTableControl>(FullPathId);
            CurrentRow = GuiTableControl.CurrentRow;
            RowCount = GuiTableControl.RowCount;
            VisibleRowCount = GuiTableControl.VisibleRowCount;
        }
        private void ResetScrolling()
        {
            Trace.WriteLine("Resetting scrollings' position to 0...");
            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.VerticalScrollbar.Position = 0;
                },
                () => { /* Do nothing */ }
            );

            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.HorizontalScrollbar.Position = 0;
                },
                () => { /* Do nothing */ }
            );
        }

        public string Info() => Info(this);
        public static string Info(Table table)
        {
            //table.Parse();
            return
                string.Join(Environment.NewLine,
                    $"Table \"{table.Name}\" captured:",
                    $"  CharHeight: \"{table.GuiTableControl.CharHeight}\"",
                    $"  CharWidth:  \"{table.GuiTableControl.CharWidth}\"",
                    $"  CharTop:    \"{table.GuiTableControl.CharTop}\"",
                    $"  CurrentCol: \"{table.GuiTableControl.CurrentCol}\"",
                    $"  CurrentRow: \"{table.GuiTableControl.CurrentRow}\"",
                    $"  RowCount:   \"{table.GuiTableControl.RowCount}\"",
                    $"  Rows.Count: \"{table.GuiTableControl.Rows.Count}\"",
                    $"  VisibleRowCount: \"{table.GuiTableControl.VisibleRowCount}\"",
                    $"  HorizontalScrollbar:",
                    $"    Minimum:  \"{table.GuiTableControl.HorizontalScrollbar.Minimum}\"",
                    $"    Maximum:  \"{table.GuiTableControl.HorizontalScrollbar.Maximum}\"",
                    $"    Position: \"{table.GuiTableControl.HorizontalScrollbar.Position}\"",
                    $"    PageSize: \"{table.GuiTableControl.HorizontalScrollbar.PageSize}\"",
                    $"  VerticalScrollBar:",
                    $"    Minimum:  \"{table.GuiTableControl.VerticalScrollbar.Minimum}\"",
                    $"    Maximum:  \"{table.GuiTableControl.VerticalScrollbar.Maximum}\"",
                    $"    Position: \"{table.GuiTableControl.VerticalScrollbar.Position}\"",
                    $"    PageSize: \"{table.GuiTableControl.VerticalScrollbar.PageSize}\"",
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
                        RefreshTableObj();

                        Trace.WriteLine(string.Join(Environment.NewLine,
                            $"Scrolling down...",
                            $"  Old position: {oldPos}",
                            $"  New position: {GuiTableControl.VerticalScrollbar.Position}"));
                        //$"  New position: {newPos}"));
                    }
                });

            RefreshTableObj(); // for the case it cannot change position
        }

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
                    RefreshTableObj();
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

            TryEmptyCellActionRelax(
                () =>
                {
                    counters["columns"] = GuiTableControl.Columns.Count;
                },
                () =>
                {
                    counters["columns"] = 0;
                });

            TryEmptyCellActionRelax(
                () =>
                {
                    counters["rows"] = GuiTableControl.RowCount;
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
            return GuiTableControl.GetCell(row % GuiTableControl.VisibleRowCount, column);
        }

        public void MakeRowVisible(int row)
        {
            int currentPage = GuiTableControl.VerticalScrollbar.Position / GuiTableControl.VisibleRowCount;

            TryEmptyCellActionRelax(
                () =>
                {
                    GuiTableControl.VerticalScrollbar.Position = row;
                });
            RefreshTableObj();
        }

        public GuiTableRow GetRow(int rowIndex)
        {
            // put the rowIndex at the top of table after scrolling
            MakeRowVisible(rowIndex);

            //Trace.WriteLine($"Will get row {rowIndex} now...");
            //return GuiTableControl.Rows.ElementAt(rowIndex % GuiTableControl.VisibleRowCount);

            // Get the top row of table
            return GuiTableControl.Rows.ElementAt(0);
        }

        public GuiTableRow GetRow(string cellValueRegex)
        {
            var foundRowIndex = _dt.Rows.Cast<DataRow>()
                .Select((row, index) => (row, index))
                .Where(
                    x => x.row.ItemArray
                        .Where(y => Rpa.IsMatch((string)y, cellValueRegex)).FirstOrDefault() != null)
                .FirstOrDefault().index;

            return GetRow(foundRowIndex);
        }

        public void SelectRow(int rowIndex) => GetRow(rowIndex).Selected = true;

        public void SelectRow(string cellValueRegex) => GetRow(cellValueRegex).Selected = true;

        public bool IsRowSelected(int rowIndex) => GetRow(rowIndex).Selected;

        public bool IsRowSelected(string cellValueRegex) => GetRow(cellValueRegex).Selected;

        public bool IsRowSelectable(int rowIndex) => GetRow(rowIndex).Selectable;

        public bool IsRowSelectable(string cellValueRegex) => GetRow(cellValueRegex).Selectable;

        public void SelectAllCols() => GuiTableControl.SelectAllColumns();

        public void DeselectAllCols() => GuiTableControl.DeselectAllColumns();

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
                    /*
                    try
                    {
                        if (!string.IsNullOrEmpty(GuiTableControl.GetCell(row, col).Text))
                        {
                            fulfilledRows++;
                            break;
                        }
                    }
                    catch(COMException ex)
                    {
                        if (ex.Message.Equals("The server threw an exception. (Exception from HRESULT: 0x80010105 (RPC_E_SERVERFAULT))"))
                        {
                            Trace.WriteLine("Caught a null empty cell. Skipping...");
                            RefreshTableObj();
                            break;
                        }
                        else
                        {
                            Trace.WriteLine("Unknown COMException occurred." +
                                " It's different from the exception that signals empty cells." +
                                $" See:\n{ex}");
                            throw ex;
                        }
                    }
                    */
                }

                // if last line is empty, stop analyzing lines
                if (fulfilledRows != row + 1)
                    break;
            }

            ResetScrolling();

            return fulfilledRows;
        }

        public bool IsEmpty()
        {
            bool isEmpty = false;

            if (FulfilledRowsCount() == 0)
                isEmpty = true;

            Trace.WriteLine($"Table \"{Name}\" is empty? {isEmpty}");

            return isEmpty;
        }

        public void Parse()
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

        public void PrintDataTable() => Trace.WriteLine(Rpa.PrintDataTable(_dt), withTimeSpec: false, color: ConsoleColor.Magenta);
        /*
                public void Find(TableAction action)
                {


                    for (long i = 0; i < GuiTableControl.RowCount; i++)
                        for (long j = 0; j < GuiTableControl.Columns.Count; j++)
                            if (action == TableAction.Select) 
                }
        */
    }
}
