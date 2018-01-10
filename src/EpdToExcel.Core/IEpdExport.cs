using EpdToExcel.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpdToExcel.Core
{
    public interface IEpdExport
    {
        void ExportEpd(IEnumerable<Epd> epds);
    }
}
