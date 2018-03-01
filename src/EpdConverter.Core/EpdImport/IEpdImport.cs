using EpdConverter.Core.Models;
using System.Collections.Generic;

namespace EpdConverter.Core.EpdImport
{
    public interface IEpdImport
    {
        /// <summary>
        /// Imports Epd datasets from a given path.
        /// </summary>
        /// <param name="path">A Filesystem path, an Url or some other path</param>
        /// <param name="indicatorFilter"></param>
        /// <returns></returns>
        IEnumerable<Epd> GetEpd(string path);
    }
}
