using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using Exc = RpaLib.Tracing.Excel;

namespace RpaLibXunit
{
    public class Excel
    {
        [Fact]
        public void ReadXlsxForceNoMacros()
        {
            var content = Exc.ReadAll(@"C:\_uvnc_transfer\901166628.xlsx", "Template para Alterações", disableMacros: true);
        }
    }
}
