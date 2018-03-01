using EpdConverter.Core.Models;
using System;
using System.Collections.Generic;
using EpdConverter.Core.EpdImport;
using EpdConverter.Core.EpdExport;

namespace EpdConverter.Core
{
    // Facade: https://en.wikipedia.org/wiki/Facade_pattern
    public static class EpdConvert
    {

        /************************************************
                      Facades public API
        ************************************************/
        public static IEnumerable<Epd> ImportEpdFromFile(
            string filePath,
            int productNumber,
            IEnumerable<string> indicatorFilter,
            Action<string> log)
        {
            var res = Import(new EpdFileImport(productNumber, indicatorFilter, log), filePath);

            return res;
        }


        public static void ExportEpdToExcel(string excelFileName, IEnumerable<IEnumerable<Epd>> epds)
        {
            Export(new EpdToXlsx(excelFileName), epds);
        }


        /************************************************
                          Privates
        ************************************************/

        private static IEnumerable<Epd> Import(IEpdImport importer, string path)
        {
            return importer.GetEpd(path);
        }

        private static void Export(IEpdExport exporter, IEnumerable<IEnumerable<Epd>> epds)
        {
            exporter.ExportEpd(epds);
        }
    }
}
