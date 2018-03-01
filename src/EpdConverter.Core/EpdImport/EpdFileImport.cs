using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml.Linq;
using EpdConverter.Core.Models;

namespace EpdConverter.Core.EpdImport
{
    public class EpdFileImport : IEpdImport
    {
        private readonly IEnumerable<string> _indicatorFilter;
        private readonly int _productNumber;
        private readonly Action<string> L;

        public EpdFileImport(int productNumber, IEnumerable<string> indicatorFilter, Action<string> log)
        {
            _productNumber = productNumber;
            _indicatorFilter = indicatorFilter ?? new List<string>();
            L = log;
        }

        public IEnumerable<Epd> GetEpd(string path)
        {
            var xml = XDocument.Load(path);

            var uuid = GetUuid(xml);
            var uri = GetUri(xml);
            var datasetBaseName = GetDataSetBaseName(xml);
            var referenceUnit = GetReferenceFlowUnit(xml);
            var referenceFlowInfo = GetReferenceFlowInfo(xml);
            var referenceFlow = GetReferenceFlowMeanAmount(xml);

            var lciResults = xml.Root
                             .Elements()
                             .Where(e => e.Name.LocalName == "exchanges" || e.Name.LocalName == "LCIAResults")
                             .Elements()
                             .Where(e => e.Elements().Any(n => n.Name.LocalName == "other"))  // Skip reference data flow
                             .Where(lci => _indicatorFilter.Contains(GetIndicatorKeyValue(lci).Item1))
                             .Select(lci =>
                              new Epd
                              {
                                  Uuid = uuid,
                                  Uri = uri,
                                  Indicator = GetIndicatorKeyValue(lci).Item2,
                                  Direction = GetDirection(lci), // Input or Output
                                  Unit = GetUnit(lci),
                                  ProductionA1ToA3 = GetEnviromentalIndicatorValueA1ToA3(lci), // A1 - A3 Special case
                                  TransportA4 = GetEnviromentalIndicatorValue(lci, "A4"),
                                  BuildingProcessA5 = GetEnviromentalIndicatorValue(lci, "A5"),
                                  UsageB1 = GetEnviromentalIndicatorValue(lci, "B1"),
                                  MaintenanceB2 = GetEnviromentalIndicatorValue(lci, "B2"),
                                  RepairB3 = GetEnviromentalIndicatorValue(lci, "B3"),
                                  ReplacementB4 = GetEnviromentalIndicatorValue(lci, "B4"),
                                  ModernizationB5 = GetEnviromentalIndicatorValue(lci, "B5"),
                                  EnergyDemandB6 = GetEnviromentalIndicatorValue(lci, "B6"),
                                  WaterDemandB7 = GetEnviromentalIndicatorValue(lci, "B7"),
                                  BreakUpC1 = GetEnviromentalIndicatorValue(lci, "C1"),
                                  TransportC2 = GetEnviromentalIndicatorValue(lci, "C2"),
                                  WasteManagementC3 = GetEnviromentalIndicatorValue(lci, "C3"),
                                  WasteDisposalC4 = GetEnviromentalIndicatorValue(lci, "C4"),
                                  ReuseAndRecoveryD = GetEnviromentalIndicatorValue(lci, "D"),
                                  DataSetBaseName = datasetBaseName,
                                  ReferenceFlow = referenceFlow,
                                  ReferenceFlowUnit = referenceUnit,
                                  ReferenceFlowInfo = referenceFlowInfo,
                                  ProductNumber = _productNumber
                              });

            return lciResults;
        }

        /************************************************
                           Privates
       ************************************************/

        private string GetDataSetBaseName(XDocument xml)
        {
            var dataSetBaseNames = xml.Root
                                      .Elements()
                                      .First(e => e.Name.LocalName == "processInformation")
                                      .Elements()
                                      .First(e => e.Name.LocalName == "dataSetInformation")
                                      .Elements()
                                      .First(e => e.Name.LocalName == "name")
                                      .Elements()
                                      .Where(e => e.Name.LocalName == "baseName");

            return GetStringValueWithLanguagefilter(dataSetBaseNames, "de");
        }

        private Guid GetUuid(XDocument xml)
        {
            var uuidString = xml.Root
                                .Elements()
                                .First(e => e.Name.LocalName == "processInformation")
                                .Elements()
                                .First(e => e.Name.LocalName == "dataSetInformation")
                                .Elements()
                                .First(e => e.Name.LocalName == "UUID")
                                .Value;

            return new Guid(uuidString);
        }

        private Uri GetUri(XDocument xml)
        {
            // Get the Uuid again: the methodes execution order should not matter meaning
            // the user of this  methodes should not be forced to first call GetUuid() and THAN use this uuid to create
            // the url.

            return new Uri("http://www.oekobaudat.de/OEKOBAU.DAT/datasetdetail/process.xhtml?uuid=" + GetUuid(xml) + "&lang=de");
        }

        private string GetReferenceFlowUnit(XDocument xml)
        {

            try
            {
                var quantitativeReference = xml.Root
                                               .Elements()
                                               .First(e => e.Name.LocalName == "processInformation")
                                               .Elements()
                                               .First(e => e.Name.LocalName == "quantitativeReference")
                                               .Elements()
                                               .First(e => e.Name.LocalName == "referenceToReferenceFlow")
                                               .Value
                                               .Trim();

                var flowRefObjectId = xml.Root
                                         .Elements()
                                         .First(e => e.Name.LocalName == "exchanges")
                                         .Elements()
                                         .First(e => e.Attribute("dataSetInternalID").Value.Trim() == quantitativeReference)
                                         .Elements()
                                         .First(e => e.Name.LocalName == "referenceToFlowDataSet")
                                         .Attribute("refObjectId")
                                         .Value
                                         .Trim();

                string flowDataSetXmlString;
                using (var client = new WebClient())
                {
                    client.Encoding = Encoding.UTF8;
                    flowDataSetXmlString = client.DownloadString(Constants.FLOW_DATASET_BASE_URI + "/flows/" + flowRefObjectId + "?format=xml");
                }

                var flowPropertiesRefObjectId = XDocument.Parse(flowDataSetXmlString)
                                                 .Root
                                                 .Elements()
                                                 .First(e => e.Name.LocalName == "flowProperties")
                                                 .Elements()
                                                 .First(e => e.Attributes().Any(a => a.Name.LocalName == "dataSetInternalID" && a.Value == "0"))
                                                 .Elements()
                                                 .First(e => e.Name.LocalName == "referenceToFlowPropertyDataSet")
                                                 .Attribute("refObjectId")
                                                 .Value
                                                 .Trim();

                string flowPropertiesXmlString;
                using (var client = new WebClient())
                {
                    client.Encoding = Encoding.UTF8;
                    flowPropertiesXmlString = client.DownloadString(Constants.FLOW_DATASET_BASE_URI + "/flowproperties/" + flowPropertiesRefObjectId + "?format=xml");
                }

                var unitGroupRefObjectId = XDocument.Parse(flowPropertiesXmlString)
                                             .Root
                                             .Elements()
                                             .First(e => e.Name.LocalName == "flowPropertiesInformation")
                                             .Elements()
                                             .First(e => e.Name.LocalName == "quantitativeReference")
                                             .Elements()
                                             .First(e => e.Name.LocalName == "referenceToReferenceUnitGroup")
                                             .Attribute("refObjectId")
                                             .Value
                                             .Trim();

                string unitGroupXmlString;
                using (var client = new WebClient())
                {
                    client.Encoding = Encoding.UTF8;
                    unitGroupXmlString = client.DownloadString(Constants.FLOW_DATASET_BASE_URI + "/unitgroups/" + unitGroupRefObjectId + " ?format=xml");
                }

                var referenceToReferenceUnit = XDocument.Parse(unitGroupXmlString)
                                                        .Root
                                                        .Elements()
                                                        .First(e => e.Name.LocalName == "unitGroupInformation")
                                                        .Elements()
                                                        .First(e => e.Name.LocalName == "quantitativeReference")
                                                        .Elements()
                                                        .First(e => e.Name.LocalName == "referenceToReferenceUnit")
                                                        .Value;

                var referenceUnitNode = XDocument.Parse(unitGroupXmlString)
                                                 .Root
                                                 .Elements()
                                                 .First(e => e.Name.LocalName == "units")
                                                 .Elements()
                                                 .First(e => e.Attributes().Any(a => a.Name == "dataSetInternalID" && a.Value == referenceToReferenceUnit))
                                                 .Elements();

                var referenceUnitName = referenceUnitNode.First(e => e.Name.LocalName == "name")
                                                         .Value;

                // Check if the meanValue is the same as the parsed meanAmount from the EPD
                var meanValue = referenceUnitNode.First(e => e.Name.LocalName == "meanValue")
                                                 .Value;

                return referenceUnitName;
            }
            catch (Exception ex)
            {
                // TODO: Log ex or just throw an Exception with the message below
                L(GetUuid(xml) + ". Fetching reference unit failed.\n" + ex.ToString());
                return string.Empty;
            }
        }

        private double GetReferenceFlowMeanAmount(XDocument xml)
        {
            var referenceToReferenceFlow = xml.Root
                                .Elements()
                                .First(e => e.Name.LocalName == "processInformation")
                                .Elements()
                                .First(e => e.Name.LocalName == "quantitativeReference")
                                .Elements()
                                .First(e => e.Name.LocalName == "referenceToReferenceFlow")
                                .Value;

            var meanAmount = xml.Root
                                .Elements()
                                .First(e => e.Name.LocalName == "exchanges")
                                .Elements()
                                .First(e => e.Attribute("dataSetInternalID").Value == referenceToReferenceFlow)
                                .Elements()
                                .First(e => e.Name.LocalName == "meanAmount")
                                .Value;

            return double.Parse(meanAmount, CultureInfo.InvariantCulture);
        }


        private string GetReferenceFlowInfo(XDocument xml)
        {
            var referenceToReferenceFlow = xml.Root
                                              .Elements()
                                              .First(e => e.Name.LocalName == "processInformation")
                                              .Elements()
                                              .First(e => e.Name.LocalName == "quantitativeReference")
                                              .Elements()
                                              .First(e => e.Name.LocalName == "referenceToReferenceFlow")
                                              .Value;

            var referenceFlowUnits = xml.Root
                                        .Elements()
                                        .First(e => e.Name.LocalName == "exchanges")
                                        .Elements()
                                        .First(e => e.Attribute("dataSetInternalID").Value == referenceToReferenceFlow)
                                        .Elements()
                                        .First(e => e.Name.LocalName == "referenceToFlowDataSet")
                                        .Elements()
                                        .Where(e => e.Name.LocalName == "shortDescription");

            return GetStringValueWithLanguagefilter(referenceFlowUnits, "de");
        }

        private Tuple<string, string> GetIndicatorKeyValue(XElement lci)
        {
            var epdNameNodeMapping = new Dictionary<string, string>
            {
                ["exchange"] = "referenceToFlowDataSet",
                ["LCIAResult"] = "referenceToLCIAMethodDataSet"
            };

            var indicators = lci.Elements()
                                .First(e => e.Name.LocalName == epdNameNodeMapping[lci.Name.LocalName]) // Should crash if the indicator name node is not declared
                                .Elements()
                                .Where(e => e.Name.LocalName == "shortDescription");

            var indicator = GetStringValueWithLanguagefilter(indicators, "de"); // not realy necessary to get the "de" entry

            var indicatorKeyArray = indicator.Split(' ')
                                             .Last()
                                             .Replace("(", string.Empty)
                                             .Replace(")", string.Empty)
                                             .ToCharArray();

            // Die Ökobaudat hat Buchstabendreher in dern Indikatornamen
            // Daher wird die Reihenfolge der Buchstaben vernachlässigt

            var indicatorKey = Constants.INDICATOR_KEY_NAME_MAPPING.Keys.Single(e => e.ToCharArray().Count() == indicatorKeyArray.Count() && Enumerable.SequenceEqual(e.ToCharArray().OrderBy(x => x), indicatorKeyArray.OrderBy(x => x)));

            return new Tuple<string, string>(indicatorKey, Constants.INDICATOR_KEY_NAME_MAPPING[indicatorKey]);
        }


        private string GetDirection(XElement lci)
        {
            return lci.Elements()
                      .FirstOrDefault(e => e.Name.LocalName == "exchangeDirection")
                      ?.Value;
        }


        private string GetUnit(XElement lci)
        {
            var units = lci.Elements()
                           .First(e => e.Name.LocalName == "other")
                           .Elements()
                           .First(e => e.Name.LocalName == "referenceToUnitGroupDataSet")
                           .Elements()
                           .Where(e => e.Name.LocalName == "shortDescription");

            return GetStringValueWithLanguagefilter(units, "de");
        }

        private double? GetEnviromentalIndicatorValue(XElement lci, string module)
        {
            var epdAmount = lci.Elements()
                                .First(e => e.Name.LocalName == "other")
                                .Elements()
                                .FirstOrDefault(e => e.Name.LocalName == "amount" && e.Attributes().First(a => a.Name.LocalName == "module").Value == module);

            if (epdAmount == null)
                return null;
            else
                return double.Parse(epdAmount.Value != string.Empty ? epdAmount.Value : "0", CultureInfo.InvariantCulture);
        }


        private double? GetEnviromentalIndicatorValueA1ToA3(XElement lci)
        {
            // A1 to A3 aggregated in one module
            var aggregated = GetEnviromentalIndicatorValue(lci, "A1-A3");

            if (aggregated != null)
                return aggregated; // Return the aggregated value

            // A1 to A3 in separated modules
            var separated = new List<double?>
            {
                GetEnviromentalIndicatorValue(lci, "A1"),
                GetEnviromentalIndicatorValue(lci, "A2"),
                GetEnviromentalIndicatorValue(lci, "A3")
            };

            if (separated.All(value => value == null))
                return null; // A1 to A3 not declared, whether aggregated nor separated

            if (separated.Any(value => value == null))
                throw new ArgumentException("A1-A3 not entirely declared\n\n" + lci.ToString()); // Not expected

            return separated.Sum();
        }


        private string GetStringValueWithLanguagefilter(IEnumerable<XElement> nodes, string preferedLanguageCode)
        {
            var preferedBaseName = nodes.FirstOrDefault(e => e.Attributes().Where(a => a.Name.LocalName == "lang" && a.Value == preferedLanguageCode).Count() > 0);

            // Check if a baseName with lang=de is declared
            return preferedBaseName?.Value ?? nodes.First().Value;
        }
    }
}
