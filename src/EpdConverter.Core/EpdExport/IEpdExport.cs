using EpdConverter.Core.Models;
using System.Collections.Generic;

namespace EpdConverter.Core.EpdExport
{
    public interface IEpdExport
    {
        /// <summary>
        /// Exports a list of EPDs. One EPD corresponds to an indicator with all it's modules.
        /// </summary>
        /// <param name="epds">EPDs to Export.</param>
        void ExportEpd(IEnumerable<IEnumerable<Epd>> epds);
    }
}
