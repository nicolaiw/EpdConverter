using EpdConverter.Core.Models;
using System.Collections.Generic;

namespace EpdConverter.Core.EpdExport
{
    public interface IEpdExport
    {
        /// <summary>
        /// Exports a list of EPD's.
        /// </summary>
        /// <param name="epds">EPD's to Export.</param>
        void ExportEpd(IEnumerable<Epd> epds);
    }
}
