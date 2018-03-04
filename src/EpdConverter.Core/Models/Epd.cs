using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpdConverter.Core.Models
{
    public class Epd
    {
        /* Stores the the corresponding indicators with it's module values */
        private Dictionary<IndicatorName, EpdIndicator> _indicators;

        /* Usefull when you want to sort a list of EPDs */
        public int ProductNumber { get; set; }

        /* The EPD's Uuid */
        public Guid Uuid { get; set; }

        /* The uri to the EPD dataset */
        public Uri Uri { get; set; }

        /* The name of the EPD */
        public string DataSetBaseName { get; set; }

        /* Additional reference flow informations */
        public string ReferenceFlowInfo { get; set; }

        /* The reference flow */
        public double ReferenceFlow { get; set; }

        /* The reference flow unit */
        public string ReferenceFlowUnit { get; set; }


        public Epd()
        {
            _indicators = new Dictionary<IndicatorName, EpdIndicator>();
        }


        public EpdIndicator this[IndicatorName indicator]
        {
            get
            {
                return _indicators.ContainsKey(indicator) ? _indicators[indicator] : null;
            }
            set
            {
                _indicators[indicator] = value;
            }
        }

        public IEnumerable<EpdIndicator> GetIndicators()
        {
            foreach (var item in _indicators.Values)
            {
                yield return item;
            }
        }
    }
}
