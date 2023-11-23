using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace RpaLib.SAP
{
    public interface ISapTabular
    {
        DataTable DataTable { get; }
        void Parse();
        string Info();
    }
}
