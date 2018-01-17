﻿using EpdToExcel.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EpdToExcel.Core
{
    // TODO: IEpdImport, IEpdExport <-- use this static class as a Facade.
    //       Write class wich implements this interfaces.
    

    public static class EpdToXlsx
    {
        private const string FLOW_DATASET_BASE_URI = "http://www.oekobaudat.de/OEKOBAU.DAT/resource";


        /************************************************
                            Public API
        ************************************************/

        public static IEnumerable<Epd> GetEpdFromXml(string epdXmlPath, int productNumber, List<string> indicatorFilter)
        {
            // Another possibility would be to use XPath instead of Linq.
            // It's a matter of taste.

            var xml = XDocument.Load(epdXmlPath);

            var meanAmount = GetReferenceFlowInfo(xml);
            var referenceFlowUnit = GetReferenceFlowUnit(xml, meanAmount);

            var lciResults = xml.Root
                             .Elements()
                             .Where(e => e.Name.LocalName == "exchanges" || e.Name.LocalName == "LCIAResults")
                             .Elements()
                             .Where(e => e.Elements().Any(n => n.Name.LocalName == "other"))  // Skip reference data flow
                             .Where(lci => indicatorFilter.Contains(GetIndicatorKeyValue(lci).Item1))
                             .Select(lci =>
                              new Epd
                              {
                                  Uuid = GetUuid(xml),
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
                                  DataSetBaseName = GetDataSetBaseName(xml),
                                  ReferenceFlow = meanAmount,
                                  ReferenceFlowUnit = referenceFlowUnit,
                                  ProductNumber = productNumber
                              });

            return lciResults.ToList();
        }


        public static void ExportEpdsToXlsx(IEnumerable<IEnumerable<Epd>> epds, string excelFileName)
        {
            // Create the file using the FileInfo object
            var file = new FileInfo(excelFileName);

            // Create the package and make sure you wrap it in a using statement
            using (var package = new ExcelPackage(file))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("EPD-Daten");

                // --------- Data and styling -------------- //

                /* Headers */
                worksheet.Cells[1, 1].Value = "Indikator";
                worksheet.Cells[1, 2].Value = "Richtung";
                worksheet.Cells[1, 3].Value = "Einheit";
                worksheet.Cells[1, 4].Value = "A1-A3";
                worksheet.Cells[1, 5].Value = "A4";
                worksheet.Cells[1, 6].Value = "A5";
                worksheet.Cells[1, 7].Value = "B1";
                worksheet.Cells[1, 8].Value = "B2";
                worksheet.Cells[1, 9].Value = "B3";
                worksheet.Cells[1, 10].Value = "B4";
                worksheet.Cells[1, 11].Value = "B5";
                worksheet.Cells[1, 12].Value = "B6";
                worksheet.Cells[1, 13].Value = "B7";
                worksheet.Cells[1, 14].Value = "C1";
                worksheet.Cells[1, 15].Value = "C2";
                worksheet.Cells[1, 16].Value = "C3";
                worksheet.Cells[1, 17].Value = "C4";
                worksheet.Cells[1, 18].Value = "D";
                worksheet.Cells[1, 19].Value = "Baustoff/Prozess";
                worksheet.Cells[1, 20].Value = "Produktnummer";
                worksheet.Cells[1, 21].Value = "Referenzfluss";
                worksheet.Cells[1, 22].Value = "Referenzfluss - Einheit";
                worksheet.Cells[1, 23].Value = "UUID";

                /* Add EPDs to Worksheet */

                var rowOffset = 0;
                for (int j = 0; j < epds.Count(); j++)
                {
                    var sortedEpds = epds.ElementAt(j).ToList();
                    sortedEpds.Sort(SortByIndicator);

                    for (int i = 0; i < sortedEpds.Count(); i++)
                    {
                        var row = i + 2 + rowOffset;

                        worksheet.Cells[row, 1].Value = epds.ElementAt(j).ElementAt(i).Indicator;
                        worksheet.Cells[row, 2].Value = epds.ElementAt(j).ElementAt(i).Direction;
                        worksheet.Cells[row, 3].Value = epds.ElementAt(j).ElementAt(i).Unit;
                        InsertValueToExcelCell(worksheet.Cells[row, 4], epds.ElementAt(j).ElementAt(i).ProductionA1ToA3);
                        InsertValueToExcelCell(worksheet.Cells[row, 5], epds.ElementAt(j).ElementAt(i).TransportA4);
                        InsertValueToExcelCell(worksheet.Cells[row, 6], epds.ElementAt(j).ElementAt(i).BuildingProcessA5);
                        InsertValueToExcelCell(worksheet.Cells[row, 7], epds.ElementAt(j).ElementAt(i).UsageB1);
                        InsertValueToExcelCell(worksheet.Cells[row, 8], epds.ElementAt(j).ElementAt(i).MaintenanceB2);
                        InsertValueToExcelCell(worksheet.Cells[row, 9], epds.ElementAt(j).ElementAt(i).RepairB3);
                        InsertValueToExcelCell(worksheet.Cells[row, 10], epds.ElementAt(j).ElementAt(i).ReplacementB4);
                        InsertValueToExcelCell(worksheet.Cells[row, 11], epds.ElementAt(j).ElementAt(i).ModernizationB5);
                        InsertValueToExcelCell(worksheet.Cells[row, 12], epds.ElementAt(j).ElementAt(i).EnergyDemandB6, Color.FromArgb(255, 102, 0));
                        InsertValueToExcelCell(worksheet.Cells[row, 13], epds.ElementAt(j).ElementAt(i).WaterDemandB7, Color.FromArgb(255, 102, 0));
                        InsertValueToExcelCell(worksheet.Cells[row, 14], epds.ElementAt(j).ElementAt(i).BreakUpC1);
                        InsertValueToExcelCell(worksheet.Cells[row, 15], epds.ElementAt(j).ElementAt(i).TransportC2);
                        InsertValueToExcelCell(worksheet.Cells[row, 16], epds.ElementAt(j).ElementAt(i).WasteManagementC3);
                        InsertValueToExcelCell(worksheet.Cells[row, 17], epds.ElementAt(j).ElementAt(i).WasteDisposalC4);
                        InsertValueToExcelCell(worksheet.Cells[row, 18], epds.ElementAt(j).ElementAt(i).ReuseAndRecoveryD);
                        worksheet.Cells[row, 19].Value = epds.ElementAt(j).ElementAt(i).DataSetBaseName;
                        worksheet.Cells[row, 20].Value = epds.ElementAt(j).ElementAt(i).ProductNumber;
                        worksheet.Cells[row, 21].Value = epds.ElementAt(j).ElementAt(i).ReferenceFlow;
                        worksheet.Cells[row, 22].Value = epds.ElementAt(j).ElementAt(i).ReferenceFlowUnit;
                        worksheet.Cells[row, 23].Value = epds.ElementAt(j).ElementAt(i).Uuid;
                    }

                    rowOffset += epds.ElementAt(j).Count() + 1;
                }

                /* Format as Table */
                using (ExcelRange range = worksheet.Cells[1, 1, rowOffset, 23])
                {
                    ExcelTable table = worksheet.Tables.Add(range, "EPD-Daten");
                    table.ShowFilter = true;
                    table.ShowHeader = true;
                    table.TableStyle = TableStyles.Medium15;
                }


                /* AutoFit */
                for (int i = 1; i <= 23; i++)
                {
                    worksheet.Column(i).AutoFit();
                }


                /* Save the file */
                package.Save();
            }
        }


        /************************************************
                            Privates
        ************************************************/

        private static string GetDataSetBaseName(XDocument xml)
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

        private static Guid GetUuid(XDocument xml)
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

        private static string GetReferenceFlowUnit(XDocument xml, double meanAmount)
        {
            //var tmp = "";
            //using (var client = new WebClient())
            //{
            //    tmp = client.DownloadString(FLOW_DATASET_BASE_URI + "/unitgroups/838aaa22-0117-11db-92e3-0800200c9a66?format=xml");
            //}

            //var tmp = xml.Root
            //                               .Elements()
            //                               .First(e => e.Name.LocalName == "quantitativeReference");

            var quantitativeReference = xml.Root
                                           .Elements()
                                           .First(e => e.Name.LocalName == "processInformation")
                                           .Elements()
                                           .First(e => e.Name.LocalName == "quantitativeReference")
                                           .Elements()
                                           .First(e => e.Name.LocalName == "referenceToReferenceFlow")
                                           .Value
                                           .Trim();

            // e.g. ../flows/0ce3c9c2-0cb4-40b7-8665-e57a9d1e48fe.xml
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

            //string flowDataUri;
            //if(referenceToFlowDataSet.Attribute("uri") != null)
            //{
            //    flowDataUri = referenceToFlowDataSet.Attribute("uri").Value.Trim().Remove(0, 2);
            //}
            //else
            //{
            //    flowDataUri = "/flows/" + referenceToFlowDataSet.Attribute("refObjectId").Value.Trim();
            //}
                                           

            string flowDataSetXmlString;
            using (var client = new WebClient())
            {
                flowDataSetXmlString = client.DownloadString(FLOW_DATASET_BASE_URI + "/flows/" + flowRefObjectId + "?format=xml");
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
                flowPropertiesXmlString = client.DownloadString(FLOW_DATASET_BASE_URI + "/flowproperties/" + flowPropertiesRefObjectId + "?format=xml");
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
                //unitGroupXmlString = client.DownloadString(FLOW_DATASET_BASE_URI + "/flowproperties/" + unitGroupUriRefObjectId + "?format=xml");
                unitGroupXmlString = client.DownloadString(FLOW_DATASET_BASE_URI +  "/unitgroups/" + unitGroupRefObjectId + " ?format=xml");

                //http://lca.jrc.ec.europa.eu/lcainfohub/datasets/ilcd/flowproperties/
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

            if(meanAmount != double.Parse(meanValue, CultureInfo.InvariantCulture))
            {
                // TODO: Include more informations (e.g. UUID)
                throw new Exception("meanAmount from EPD != meanValue from referenceUnit");
            }

            return referenceUnitName;
        }

        private static double GetReferenceFlowInfo(XDocument xml)
        {
            var meanAmount = xml.Root
                                .Elements()
                                .First(e => e.Name.LocalName == "exchanges")
                                .Elements()
                                .First(e => e.Elements().Where(i => i.Name.LocalName == "meanAmount").Count() == 1)
                                .Elements()
                                .First(e => e.Name.LocalName == "meanAmount")
                                .Value;

            return double.Parse(meanAmount, CultureInfo.InvariantCulture);
        }


        //private static string GetReferenceFlowUnit(XDocument xml)
        //{
        //    var referenceFlowUnits = xml.Root
        //                               .Elements()
        //                               .First(e => e.Name.LocalName == "exchanges")
        //                               .Elements()
        //                               .First(e => e.Elements().Where(i => i.Name.LocalName == "meanAmount").Count() == 1)
        //                               .Elements()
        //                               .First(e => e.Name.LocalName == "referenceToFlowDataSet")
        //                               .Elements()
        //                               .Where(e => e.Name.LocalName == "shortDescription");

        //    return GetStringValueWithLanguagefilter(referenceFlowUnits, "de");
        //}


        private static void InsertValueToExcelCell(ExcelRange range, double? value, Color? color = null)
        {
            range.Style.Fill.PatternType = ExcelFillStyle.Solid;

            if (value.HasValue)
            {
                range.Value = value;
                range.Style.Fill.BackgroundColor.SetColor(color ?? Color.FromArgb(0, 255, 0));
            }
            else
            {
                range.Value = 0;
                range.Style.Fill.BackgroundColor.SetColor(color ?? Color.FromArgb(255, 0, 0));
            }
        }

        public static Dictionary<string, string> IndicatorKeyNameMapping = new Dictionary<string, string>
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

        private static int SortByIndicator(Epd a, Epd b)
        {
            var indicatorAcronymA = a.Indicator.Split(' ').Last().Replace("(", string.Empty).Replace(")", string.Empty);
            var indicatorAcronymB = b.Indicator.Split(' ').Last().Replace("(", string.Empty).Replace(")", string.Empty);

            var indexOfA = Array.IndexOf(IndicatorKeyNameMapping.Keys.ToArray(), indicatorAcronymA);
            var indexOfB = Array.IndexOf(IndicatorKeyNameMapping.Keys.ToArray(), indicatorAcronymB);

            return indexOfA < indexOfB ? -1 : 1;
        }

        private static Tuple<string, string> GetIndicatorKeyValue(XElement lci)
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

            var indicatorKey = IndicatorKeyNameMapping.Keys.Single(e => e.ToCharArray().Count() == indicatorKeyArray.Count() && Enumerable.SequenceEqual(e.ToCharArray().OrderBy(x => x), indicatorKeyArray.OrderBy(x => x)));

            return new Tuple<string,string>(indicatorKey, IndicatorKeyNameMapping[indicatorKey]);
        }


        private static string GetDirection(XElement lci)
        {
            return lci.Elements()
                      .FirstOrDefault(e => e.Name.LocalName == "exchangeDirection")
                      ?.Value;
        }


        private static string GetUnit(XElement lci)
        {
            var units = lci.Elements()
                           .First(e => e.Name.LocalName == "other")
                           .Elements()
                           .First(e => e.Name.LocalName == "referenceToUnitGroupDataSet")
                           .Elements()
                           .Where(e => e.Name.LocalName == "shortDescription");

            return GetStringValueWithLanguagefilter(units, "de");
        }

        private static double? GetEnviromentalIndicatorValue(XElement lci, string module)
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


        private static double? GetEnviromentalIndicatorValueA1ToA3(XElement lci)
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


        private static string GetStringValueWithLanguagefilter(IEnumerable<XElement> nodes, string preferedLanguageCode)
        {
            var preferedBaseName = nodes.FirstOrDefault(e => e.Attributes().Where(a => a.Name.LocalName == "lang" && a.Value == preferedLanguageCode).Count() > 0);

            // Check if a baseName with lang=de is declared
            return preferedBaseName?.Value ?? nodes.First().Value;
        }

    }
}
