using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.Tracing.Exceptions
{
    public class WorksheetNotFoundException : ExcelException
    {
        public WorksheetNotFoundException(string sheetName, string fileName) : base(GetMsg(sheetName, fileName)) { }

        private static string GetMsg(string sheetName, string fileName)
        {
            return $"The worksheet \"{sheetName}\" wasn't found within file: \"{fileName}\".";
        }
    }
}
