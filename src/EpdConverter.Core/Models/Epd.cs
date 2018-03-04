using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpdConverter.Core.Models
{
    public class Epd
    {
        private Dictionary<IndicatorName, EpdIndicator> _indicators;
        
        public int ProductNumber { get; set; } /* Usefull when you want to sort a list of EPDs */
        public Guid Uuid { get; set; }
        public Uri Uri { get; set; }
        public string DataSetBaseName { get; set; }
        public string ReferenceFlowInfo { get; set; }
        public double ReferenceFlow { get; set; }
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
