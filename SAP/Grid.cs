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
    public class Grid : SapComponent, ISapTabular
    {
        public GuiGridView GuiGridView { get; private set; }
        public DataTable DataTable { get; private set; }

        public Grid(Session session, GuiGridView guiGridView) : base(session)
        {
            GuiGridView = guiGridView;
            FullPathId = Rpa.Replace(guiGridView.Id, @"/app/con\[\d+\]/ses\[\d+\]/", string.Empty);
            Refresh();
        }

        public Grid(Session session,  string fullPathId) : this(session, session.FindById<GuiGridView>(fullPathId))
        { }

        // refreshes datatable info after updating the GuiGridView object
        public void Refresh()
        {
            DataTable = new DataTable();
            GuiGridView = Session.FindById<GuiGridView>(FullPathId);
            Parse();
            Log.Write(Info());
        }

        /// <summary>
        /// Generate the DataTable property by parsing the GuiGridView object.
        /// </summary>
        public void Parse()
        {
            GuiGridView.SelectAll();
            //Cols = _grid.SelectedColumns;
            //Rows = long.Parse(Regex.Match(_grid.SelectedRows, @"(?<=\d+-)\d+").Value) + 1;

            foreach (var col in GuiGridView.SelectedColumns)
            {
                DataTable.Columns.Add(col, typeof(string));
            }

            for (int i = 0; i < GuiGridView.RowCount; i++)
            {
                DataRow datarow = DataTable.NewRow();

                if (i != 0 && i % GuiGridView.VisibleRowCount == 0)
                {
                    GuiGridView.FirstVisibleRow = i;
                }

                foreach (DataColumn col in DataTable.Columns)
                {
                    //_grid.SetCurrentCell(i, col);
                    datarow[col.ColumnName] = GuiGridView.GetCellValue(i, col.ColumnName);
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
            string[] selectedColumns = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.SelectedColumns).ToArray();
            string[] selectedCells = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.SelectedCells).ToArray();
            string[] columnOrder = Rpa.COMCollectionToICollection<string>(grid.GuiGridView.ColumnOrder).ToArray();

            return
              string.Join(Environment.NewLine,
                  $"Grid \"{grid.Name}\" captured:",
                  $"  ColumnCount:         \"{grid.GuiGridView.ColumnCount}\"",
                  $"  ColumnOrder :        \"{string.Join(", ", columnOrder)}\"",
                  $"  CurrentCellColumn :  \"{grid.GuiGridView.CurrentCellColumn}\"",
                  $"  CurrentCellRow :     \"{grid.GuiGridView.CurrentCellRow}\"",
                  $"  FirstVisibleColumn : \"{grid.GuiGridView.FirstVisibleColumn}\"",
                  $"  FirstVisibleRow :    \"{grid.GuiGridView.FirstVisibleRow}\"",
                  $"  FrozenColumnCount :  \"{grid.GuiGridView.FrozenColumnCount}\"",
                  $"  RowCount:            \"{grid.GuiGridView.RowCount}\"",
                  $"  SelectedCells:       \"{string.Join(", ", selectedCells)}\"",
                  $"  SelectedColumns:     \"{string.Join(", ", selectedColumns)}\"",
                  $"  SelectedRows:        \"{grid.GuiGridView.SelectedRows}\"",
                  $"  SelectionMode:       \"{grid.GuiGridView.SelectionMode}\"",
                  $"  Title:               \"{grid.GuiGridView.Title}\"",
                  $"  ToolbarButtonCount:  \"{grid.GuiGridView.ToolbarButtonCount}\"",
                  $"  VisibleRowCount:     \"{grid.GuiGridView.VisibleRowCount}\"",
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
