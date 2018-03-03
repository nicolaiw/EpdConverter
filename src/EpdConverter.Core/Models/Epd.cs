using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EpdConverter.Core.Models
{
    public class Epd
    {
        /* Usefull when you want to sort a list of EPDs */
        public int ProductNumber { get; private set; }

        private Dictionary<IndicatorName, EpdIndicator> _indicators;

        public Epd()
        {
            _indicators = new Dictionary<IndicatorName, EpdIndicator>();
        }

        public Epd(int productNumber) : this()
        {
            ProductNumber = productNumber;
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
