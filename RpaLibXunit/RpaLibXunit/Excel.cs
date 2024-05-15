using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Exc = RpaLib.Tracing.Excel;
using RpaLib.ProcessAutomation;
using RpaLib.Tracing;

namespace RpaLibXunit
{
    public class Excel
    {
        [Fact]
        public void ReadXlsxForceNoMacros()
        {
            var content = Exc.ReadAll(@"C:\_uvnc_transfer\901166628.xlsx", "Template para Alterações", disableMacros: true);
        }

        [Fact]
        public void ReadAllMutiProc()
        {
            var content = Exc.ReadAll(@"%USERPROFILE%\Downloads\MesCronogramaFechamento\04 - ABRIL\Messer_R2R_Cronograma Fechamento_abril.2024.xlsx", @"CloseTasks(C)", startRow: 8);
        }

        [Fact]
        public void InsertFormula()
        {
            var excelFile = @"%USERPROFILE%\Downloads\Test.xlsx";
            var aba = @"Sheet1";
            var formula = @"=VLOOKUP(G2,A:C,3,0)";

            //Exc.InsertFormula(excelFile, aba, "H1:H20", formula, visible: true);
            Exc.InsertFormula(excelFile, aba, formula, 1, 8, lastCellCol: 8, visible: true);
        }

        [Fact]
        public void ReadAllSheets()
        {
            var excelFile = @"%USERPROFILE%\Downloads\PlanilhaAtividades\CronogramaFechamento_Datas_e_Responsáveis.xlsx";

            var sheetsToRead = new SheetInfo[]
            {
                new SheetInfo()
                {
                    FirstCol = ExcelColumn.F,
                    LastCol = ExcelColumn.H
                },
                new SheetInfo()
            };

            var sheets = Exc.ReadAllSheets(excelFile, sheetsToRead, visible: true);
        }

        [Fact]
        public void DeleteFromRange()
        {
            var excelFile = @"%USERPROFILE%\Downloads\PlanilhaAtividades\CronogramaFechamento_Datas_e_Responsáveis.xlsx";
            var sheetName = "Datas";

            var excel = new Exc(excelFile, sheetName);
            excel.ToggleVisible(true);

            var firstRow = 2;
            var usedRange = excel.GetUsedRange($"{ExcelColumn.F}:{ExcelColumn.H}");

            excel.DeleteCells($"{ExcelColumn.F}{firstRow}:{ExcelColumn.H}{usedRange.LastRow}");

            excel.SaveAndQuit();
        }
    }
}
