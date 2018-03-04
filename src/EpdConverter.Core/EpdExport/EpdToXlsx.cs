using EpdConverter.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace EpdConverter.Core.EpdExport
{
    public class EpdToXlsx : IEpdExport
    {
        private readonly string _excelFileName;

        public EpdToXlsx(string excelFileName)
        {
            _excelFileName = excelFileName;
        }

        public void ExportEpd(IEnumerable<Epd> epds)
        {
            // Create the file using the FileInfo object
            var file = new FileInfo(_excelFileName);

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
                worksheet.Cells[1, 23].Value = "Referenzfluss - Info";
                worksheet.Cells[1, 24].Value = "UUID";
                worksheet.Cells[1, 25].Value = "Ökobaudat - Link";

                /* Add EPDs to Worksheet */

                var rowOffset = 0;
                for (int j = 0; j < epds.Count(); j++)
                {
                    var sortedEpdIndicators = epds.ElementAt(j).GetIndicators().ToList();
                    sortedEpdIndicators.Sort(SortByIndicator);

                    for (int i = 0; i < sortedEpdIndicators.Count(); i++)
                    {
                        var row = i + 2 + rowOffset;

                        worksheet.Cells[row, 1].Value = sortedEpdIndicators[i].IndicatorDescription;
                        worksheet.Cells[row, 2].Value = sortedEpdIndicators[i].Direction;
                        worksheet.Cells[row, 3].Value = sortedEpdIndicators[i].Unit;
                        InsertValueToExcelCell(worksheet.Cells[row, 4], sortedEpdIndicators[i].ProductionA1ToA3);
                        InsertValueToExcelCell(worksheet.Cells[row, 5], sortedEpdIndicators[i].TransportA4);
                        InsertValueToExcelCell(worksheet.Cells[row, 6], sortedEpdIndicators[i].BuildingProcessA5);
                        InsertValueToExcelCell(worksheet.Cells[row, 7], sortedEpdIndicators[i].UsageB1);
                        InsertValueToExcelCell(worksheet.Cells[row, 8], sortedEpdIndicators[i].MaintenanceB2);
                        InsertValueToExcelCell(worksheet.Cells[row, 9], sortedEpdIndicators[i].RepairB3);
                        InsertValueToExcelCell(worksheet.Cells[row, 10], sortedEpdIndicators[i].ReplacementB4);
                        InsertValueToExcelCell(worksheet.Cells[row, 11], sortedEpdIndicators[i].ModernizationB5);
                        InsertValueToExcelCell(worksheet.Cells[row, 12], sortedEpdIndicators[i].EnergyDemandB6, Color.FromArgb(255, 102, 0));
                        InsertValueToExcelCell(worksheet.Cells[row, 13], sortedEpdIndicators[i].WaterDemandB7, Color.FromArgb(255, 102, 0));
                        InsertValueToExcelCell(worksheet.Cells[row, 14], sortedEpdIndicators[i].BreakUpC1);
                        InsertValueToExcelCell(worksheet.Cells[row, 15], sortedEpdIndicators[i].TransportC2);
                        InsertValueToExcelCell(worksheet.Cells[row, 16], sortedEpdIndicators[i].WasteManagementC3);
                        InsertValueToExcelCell(worksheet.Cells[row, 17], sortedEpdIndicators[i].WasteDisposalC4);
                        InsertValueToExcelCell(worksheet.Cells[row, 18], sortedEpdIndicators[i].ReuseAndRecoveryD);
                        worksheet.Cells[row, 19].Value = epds.ElementAt(j).DataSetBaseName;
                        worksheet.Cells[row, 20].Value = epds.ElementAt(j).ProductNumber;
                        worksheet.Cells[row, 21].Value = epds.ElementAt(j).ReferenceFlow;
                        worksheet.Cells[row, 22].Value = epds.ElementAt(j).ReferenceFlowUnit;
                        worksheet.Cells[row, 23].Value = epds.ElementAt(j).ReferenceFlowInfo;
                        worksheet.Cells[row, 24].Value = epds.ElementAt(j).Uuid;
                        worksheet.Cells[row, 25].Hyperlink = epds.ElementAt(j).Uri;
                        worksheet.Cells[row, 25].Value = "Link zur EPD";
                    }

                    rowOffset += sortedEpdIndicators.Count() + 1;
                }

                /* Format as Table */
                using (ExcelRange range = worksheet.Cells[1, 1, rowOffset, 25])
                {
                    ExcelTable table = worksheet.Tables.Add(range, "EPD-Daten");
                    table.ShowFilter = true;
                    table.ShowHeader = true;
                    table.TableStyle = TableStyles.Medium15;
                }


                /* AutoFit */
                for (int i = 1; i <= 25; i++)
                {
                    worksheet.Column(i).AutoFit();
                }

                /* Save the file */
                package.Save();
            }
        }

        private void InsertValueToExcelCell(ExcelRange range, double? value, Color? color = null)
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

        private int SortByIndicator(EpdIndicator a, EpdIndicator b)
        {
            var indicatorAcronymA = a.IndicatorDescription.Split(' ').Last().Replace("(", string.Empty).Replace(")", string.Empty);
            var indicatorAcronymB = b.IndicatorDescription.Split(' ').Last().Replace("(", string.Empty).Replace(")", string.Empty);

            var indexOfA = Array.IndexOf(Constants.INDICATOR_KEY_NAME_MAPPING.Keys.ToArray(), indicatorAcronymA);
            var indexOfB = Array.IndexOf(Constants.INDICATOR_KEY_NAME_MAPPING.Keys.ToArray(), indicatorAcronymB);

            return indexOfA < indexOfB ? -1 : 1;
        }
      
    }
}
