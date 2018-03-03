using EpdConverter.Core.Models;
using System.Collections.Generic;

namespace EpdConverter.Core.EpdImport
{
    public interface IEpdImport
    {
        /// <summary>
        /// Imports an EPD dataset from a given path.
        /// </summary>
        /// <param name="path">A filesystem path, an Url or some other path.</param>
        /// <returns>List of EPD's whereas an EPD corresponds to an indicator with all its modules.</returns>
        Epd GetEpd(string path);
    }
}
