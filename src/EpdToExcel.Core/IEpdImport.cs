using EpdToExcel.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpdToExcel.Core
{
    public interface IEpdImport
    {
        /// <summary>
        /// Imports Epd datasets from a given path.
        /// </summary>
        /// <param name="path">A Filesystem path, an Url or some other path</param>
        /// <returns></returns>
        IEnumerable<Epd> GetEpd(string path);
    }
}
