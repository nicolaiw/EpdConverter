using EpdConverter.Core.Models;
using System.Collections.Generic;

namespace EpdConverter.Core.EpdExport
{
    public interface IEpdExport
    {
        /// <summary>
        /// Exports a list of Epd's. One Epd corresponds to an indicator wiht all it's modules.
        /// </summary>
        /// <param name="epds">Epd's to Export</param>
        void ExportEpd(IEnumerable<IEnumerable<Epd>> epds);
    }
}
