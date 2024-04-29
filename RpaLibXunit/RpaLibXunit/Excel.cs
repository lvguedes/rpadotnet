using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Exc = RpaLib.Tracing.Excel;
using RpaLib.ProcessAutomation;

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
    }
}
