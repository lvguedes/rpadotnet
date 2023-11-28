using RpaLib.ProcessAutomation;
using sapfewse;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Tracing;

namespace RpaLib.SAP
{
    // Prefix: wnd[N]/usr/cntlGRID1/shellcont/shell/shellcont[N]/shell
    public class Grid : SapComponent<GuiGridView>, ISapTabular
    {
        public SapComWrapper<GuiGridView> GuiGridView { get; private set; }
        public DataTable DataTable { get; private set; }

        public Grid(Session session, GuiGridView guiGridView) : base(session, guiGridView)
        {
            GuiGridView = new SapComWrapper<GuiGridView>(guiGridView);
            Refresh();
        }

        public Grid(Session session,  string fullPathId) : this(session, (GuiGridView)session.FindById<GuiGridView>(fullPathId).Com)
        { }

        // refreshes datatable info after updating the GuiGridView object
        public void Refresh()
        {
            DataTable = new DataTable();
            GuiGridView = Session.FindById<GuiGridView>(PathId);
            Parse();
            Log.Write(Info());
        }

        /// <summary>
        /// Generate the DataTable property by parsing the GuiGridView object.
        /// </summary>
        public void Parse()
        {
            GuiGridView.Com.SelectAll(); // auto-paginate the full table
            //Cols = _grid.SelectedColumns;
            //Rows = long.Parse(Regex.Match(_grid.SelectedRows, @"(?<=\d+-)\d+").Value) + 1;

            foreach (var col in GuiGridView.Com.SelectedColumns as dynamic)
            {
                DataTable.Columns.Add(col, typeof(string));
            }

            for (int i = 0; i < GuiGridView.Com.RowCount; i++)
            {
                DataRow datarow = DataTable.NewRow();

                if (i != 0 && i % GuiGridView.Com.VisibleRowCount == 0)
                {
                    GuiGridView.Com.FirstVisibleRow = i;
                }

                foreach (DataColumn col in DataTable.Columns)
                {
                    //_grid.SetCurrentCell(i, col);
                    datarow[col.ColumnName] = GuiGridView.Com.GetCellValue(i, col.ColumnName);
                }
                DataTable.Rows.Add(datarow);
            }
        }

        public bool IsEmpty()
        {
            if (DataTable.Rows.Count == 0)
                return true;
            else
                return false;
        }

        public int FulfilledRowsCount() => DataTable.Rows.Count;


        public string Info() => Info(this);
        public static string Info(Grid grid)
        {
            string[] columnNamesDatatable = grid.DataTable.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
            string[] selectedColumns = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.Com.SelectedColumns).ToArray();
            string[] selectedCells = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.Com.SelectedCells).ToArray();
            string[] columnOrder = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.Com.ColumnOrder).ToArray();

            return
              string.Join(Environment.NewLine,
                  $"Grid \"{grid.Name}\" captured:",
                  $"  ColumnCount:         \"{grid.GuiGridView.Com.ColumnCount}\"",
                  $"  ColumnOrder :        \"{string.Join(", ", columnOrder)}\"",
                  $"  CurrentCellColumn :  \"{grid.GuiGridView.Com.CurrentCellColumn}\"",
                  $"  CurrentCellRow :     \"{grid.GuiGridView.Com.CurrentCellRow}\"",
                  $"  FirstVisibleColumn : \"{grid.GuiGridView.Com.FirstVisibleColumn}\"",
                  $"  FirstVisibleRow :    \"{grid.GuiGridView.Com.FirstVisibleRow}\"",
                  $"  FrozenColumnCount :  \"{grid.GuiGridView.Com.FrozenColumnCount}\"",
                  $"  RowCount:            \"{grid.GuiGridView.Com.RowCount}\"",
                  $"  SelectedCells:       \"{string.Join(", ", selectedCells)}\"",
                  $"  SelectedColumns:     \"{string.Join(", ", selectedColumns)}\"",
                  $"  SelectedRows:        \"{grid.GuiGridView.Com.SelectedRows}\"",
                  $"  SelectionMode:       \"{grid.GuiGridView.Com.SelectionMode}\"",
                  $"  Title:               \"{grid.GuiGridView.Com.Title}\"",
                  $"  ToolbarButtonCount:  \"{grid.GuiGridView.Com.ToolbarButtonCount}\"",
                  $"  VisibleRowCount:     \"{grid.GuiGridView.Com.VisibleRowCount}\"",
                  $"  Grid.DataTable.Columns.Count: {string.Join(", ", columnNamesDatatable)}",
                  $"  Grid.DataTable.Rows.Count: \"{grid.DataTable.Rows.Count}\""
                  );
        }

        public string PrintDataTable() => Rpa.PrintDataTable(DataTable);

        /*
        TODO: 
        Methods to check:
        string GetCellState(long Row, string Column) {Possible return: Normal, Error, Warning, Info}
        string GetCellType(long Row, string Column) {Possible return: Normal, Button, Checkbox, ValueList, RadioButton}
        string GetCellValue(long Row, string Column)
        ICollection<string> GetColumnTitles(string Column) // all possibilities of given column titles
        string GetDisplayedColumnTitle(string Column) // one of the values returned by function above
        void SetCurrentCell(long Row, string Column)

        void SelectAll() // selects all columns and rows
        ICollection<string> SelectedColumns { get; set; } // Setting this property can raise an exception, if the new collection contains an invalid column identifier.

        Properties:
        long CurrentCellRow { get; set; }
        string CurrentCellColumn { get; set; }
         */
    }
}
