using EpdConverter.Core.Models;
using System.Collections.Generic;

namespace EpdConverter.Core.EpdImport
{
    public interface IEpdImport
    {
        /// <summary>
        /// Imports an EPD datasets from a given path.
        /// </summary>
        /// <param name="path">A Filesystem path, an Url or some other path.</param>
        /// <returns>List of EPDs wheres an EPD corresponds to an indicator with all its modules.</returns>
        IEnumerable<Epd> GetEpd(string path);
    }
}
