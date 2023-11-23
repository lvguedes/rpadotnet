using RpaLib.ProcessAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.SAP.Model;
using RpaLib.Tracing;
using sapfewse;
using System.Data;
using RpaLib.SAP.Exceptions;

namespace RpaLib.SAP
{

    // Usually you must have a lot of GuiLabels structured directly inside
    // a GuiUserArea object
    //    /app/conn[N]/ses[N]/usr/lbl[N,N]
    // It's impossible to cast this object to GuiTableControl or GuiGridView.
    // It's only accepted as a GuiUserArea object, but this one can't be read as a table natively.
    // That's why this class exists.
    public class LabelTable
    {
        public string PathId { get; }
        public Session Session { get; }
        public GuiUserArea ParentGuiUserArea
        {
            get => Session.FindById<GuiUserArea>(PathId, suppressTrace: true);
        }
        public GuiLabel[] GuiLabels { get; private set; }
        public Label[][] TableArray { get; private set; }
        public DataTable DataTable { get; private set; }
        public Scroll VerticalScrollbar
        {
            get => new Scroll(ParentGuiUserArea.VerticalScrollbar);
        }
        public Scroll HorizontalScrollbar
        {
            get => new Scroll(ParentGuiUserArea.HorizontalScrollbar);
        }

        public LabelTable(Session session, string guiUserAreaPathId)
        {
            PathId = guiUserAreaPathId;
            Session = session;

            GuiLabels = Sap.FindByType<GuiLabel>(ParentGuiUserArea);

            ParseTable();
            GetDataTable();
        }


        // TODO: Add pagination using vertical and horizontal scrolls
        private Label[][] ParseTable(List<Label[]> tableList = null)
        {
            const string colRegex = @"(?<=lbl\[)\d+(?=,\s*\d+\]$)";
            const string rowRegex = @"(?<=lbl\[\d+,\s*)\d+(?=\]$)";

            HashSet<int> rows = new HashSet<int>();
            HashSet<int> columns = new HashSet<int>();
            List<Label> labels = new List<Label>();

            foreach (var guiLabel in GuiLabels)
            {
                var currentLabel = new Label
                {
                    Row = int.Parse(Rpa.Match(guiLabel.Id, rowRegex)),
                    Col = int.Parse(Rpa.Match(guiLabel.Id, colRegex)),
                    Text = guiLabel.Text,
                    GuiLabel = guiLabel,
                };

                labels.Add(currentLabel);
                rows.Add(currentLabel.Row);
                columns.Add(currentLabel.Col);
            }

            var orderedRows = rows.Cast<int>().OrderBy(x => x).ToArray();
            var orderedCols = columns.Cast<int>().OrderBy(x => x).ToArray();

            tableList = tableList ?? new List<Label[]>();
            foreach (var nRow in orderedRows)
            {
                Label[] rowList = labels.Where(x => x.Row == nRow).OrderBy(x => x.Col).ToArray();
                tableList.Add(rowList);
            }

            // Scroll down if needed
            if (VerticalScrollbar.IsNeeded())
            {
                VerticalScrollbar.NextPage();
                ParseTable(tableList);
            }
            else
            {
                TableArray = tableList.ToArray();
            }

            return TableArray;
        }

        private DataTable GetDataTable(int header = 0)
        {
            DataTable = new DataTable();

            // add the header row (DataColumns) to dataTable (create)
            foreach (var label in TableArray[header])
            {
                var column = new DataColumn
                {
                    DataType = typeof(string),
                    ColumnName = label.Text,
                };

                DataTable.Columns.Add(column);
            }

            // add the rows to datatable (populate)
            for (var i = 0; i < TableArray.Length; i++)
            {
                // skip the header row
                if (i == header)
                    continue;

                var row = DataTable.NewRow();

                for (var j = 0; j < TableArray[i].Length; j++)
                {
                    row[j] = TableArray[i][j].Text;
                }

                DataTable.Rows.Add(row);
            }

            return DataTable;
        }

        public void SelectCell(int row, int column)
        {
            TableArray[row][column].GuiLabel.SetFocus();
        }

        public void SelectCell(string textRegex, int defaultToSelect = 0)
        {
            var guiLabelsFound = TableArray.Select(x => x.Where(y => Rpa.IsMatch(y.Text, textRegex))).Cast<Label>().ToArray();
            Label cell;

            if (guiLabelsFound.Length == 0)
            {
                throw new GuiLabelTextRegexException(textRegex);
            }
            else if (guiLabelsFound.Length > 1)
            {
                cell = guiLabelsFound[defaultToSelect];
            }
            else
            {
                cell = guiLabelsFound[0];
            }

            cell.GuiLabel.SetFocus();
        }

        public void PrintDataTable()
        {
            Trace.WriteLine($"Printing the DataTable that represents SAP GuiLabel Table \"{ParentGuiUserArea.Id}\":", color: ConsoleColor.Yellow);
            Trace.WriteLine(Rpa.PrintDataTable(DataTable), withTimeSpec: false, color: ConsoleColor.Magenta);
        }
    }
}
