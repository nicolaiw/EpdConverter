using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpdConverter.Core
{
    public static class Constants
    {
        /// <summary>
        /// Ökobaudat base url.
        /// </summary>
        public const string FLOW_DATASET_BASE_URI = "http://www.oekobaudat.de/OEKOBAU.DAT/resource";

        /// <summary>
        /// ILCA Indicators.
        /// </summary>
        public static readonly Dictionary<string, string> INDICATOR_KEY_NAME_MAPPING = new Dictionary<string, string>
        {
            ["PERE"] = "Erneuerbare Primärenergie als Energieträger (PERE)",
            ["PERM"] = "Erneuerbare Primärenergie zur stofflichen Nutzung (PERM)",
            ["PERT"] = "Total erneuerbare Primärenergie (PERT)",
            ["PENRE"] = "Nicht-erneuerbare Primärenergie als Energieträger (PENRE)",
            ["PENRM"] = "Nicht-erneuerbare Primärenergie zur stofflichen Nutzung (PENRM)",
            ["PENRT"] = "Total nicht erneuerbare Primärenergie (PENRT)",
            ["SM"] = "Einsatz von Sekundärstoffen (SM)",
            ["RSF"] = "Erneuerbare Sekundärbrennstoffe (RSF)",
            ["NRSF"] = "Nicht erneuerbare Sekundärbrennstoffe (NRSF)",
            ["FW"] = "Einsatz von Süßwasserressourcen (FW)",
            ["HWD"] = "Gefährlicher Abfall zur Deponie (HWD)",
            ["NHWD"] = "Entsorgter nicht gefährlicher Abfall (NHWD)",
            ["RWD"] = "Entsorgter radioaktiver Abfall (RWD)",
            ["CRU"] = "Komponenten für die Wiederverwendung (CRU)",
            ["MFR"] = "Stoffe zum Recycling (MFR)",
            ["MER"] = "Stoffe für die Energierückgewinnung (MER)",
            ["EEE"] = "Exportierte elektrische Energie (EEE)",
            ["EET"] = "Exportierte thermische Energie (EET)",
            ["GWP"] = "Globales Erwärmungspotenzial (GWP)",
            ["ODP"] = "Abbaupotenzial der stratosphärischen Ozonschicht (ODP)",
            ["POCP"] = "Bildungspotenzial für troposphärisches Ozon (POCP)",
            ["AP"] = "Versauerungspotenzial (AP)",
            ["EP"] = "Eutrophierungspotenzial (EP)",
            ["ADPE"] = "Potenzial für den abiotischen Abbau nicht fossiler Ressourcen (ADPE)",
            ["ADPF"] = "Potenzial für den abiotischen Abbau fossiler Brennstoffe (ADPF)"
        };
    }
}
