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

        public void ExportEpd(IEnumerable<IEnumerable<Epd>> epds)
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
                    var sortedEpds = epds.ElementAt(j).ToList();
                    sortedEpds.Sort(SortByIndicator);

                    for (int i = 0; i < sortedEpds.Count(); i++)
                    {
                        var row = i + 2 + rowOffset;

                        worksheet.Cells[row, 1].Value = sortedEpds[i].Indicator;
                        worksheet.Cells[row, 2].Value = sortedEpds[i].Direction;
                        worksheet.Cells[row, 3].Value = sortedEpds[i].Unit;
                        InsertValueToExcelCell(worksheet.Cells[row, 4], sortedEpds[i].ProductionA1ToA3);
                        InsertValueToExcelCell(worksheet.Cells[row, 5], sortedEpds[i].TransportA4);
                        InsertValueToExcelCell(worksheet.Cells[row, 6], sortedEpds[i].BuildingProcessA5);
                        InsertValueToExcelCell(worksheet.Cells[row, 7], sortedEpds[i].UsageB1);
                        InsertValueToExcelCell(worksheet.Cells[row, 8], sortedEpds[i].MaintenanceB2);
                        InsertValueToExcelCell(worksheet.Cells[row, 9], sortedEpds[i].RepairB3);
                        InsertValueToExcelCell(worksheet.Cells[row, 10], sortedEpds[i].ReplacementB4);
                        InsertValueToExcelCell(worksheet.Cells[row, 11], sortedEpds[i].ModernizationB5);
                        InsertValueToExcelCell(worksheet.Cells[row, 12], sortedEpds[i].EnergyDemandB6, Color.FromArgb(255, 102, 0));
                        InsertValueToExcelCell(worksheet.Cells[row, 13], sortedEpds[i].WaterDemandB7, Color.FromArgb(255, 102, 0));
                        InsertValueToExcelCell(worksheet.Cells[row, 14], sortedEpds[i].BreakUpC1);
                        InsertValueToExcelCell(worksheet.Cells[row, 15], sortedEpds[i].TransportC2);
                        InsertValueToExcelCell(worksheet.Cells[row, 16], sortedEpds[i].WasteManagementC3);
                        InsertValueToExcelCell(worksheet.Cells[row, 17], sortedEpds[i].WasteDisposalC4);
                        InsertValueToExcelCell(worksheet.Cells[row, 18], sortedEpds[i].ReuseAndRecoveryD);
                        worksheet.Cells[row, 19].Value = sortedEpds[i].DataSetBaseName;
                        worksheet.Cells[row, 20].Value = sortedEpds[i].ProductNumber;
                        worksheet.Cells[row, 21].Value = sortedEpds[i].ReferenceFlow;
                        worksheet.Cells[row, 22].Value = sortedEpds[i].ReferenceFlowUnit;
                        worksheet.Cells[row, 23].Value = sortedEpds[i].ReferenceFlowInfo;
                        worksheet.Cells[row, 24].Value = sortedEpds[i].Uuid;
                        worksheet.Cells[row, 25].Hyperlink = sortedEpds[i].Uri;
                        worksheet.Cells[row, 25].Value = "Link zur EPD";
                    }

                    rowOffset += sortedEpds.Count() + 1;
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

        private int SortByIndicator(Epd a, Epd b)
        {
            var indicatorAcronymA = a.Indicator.Split(' ').Last().Replace("(", string.Empty).Replace(")", string.Empty);
            var indicatorAcronymB = b.Indicator.Split(' ').Last().Replace("(", string.Empty).Replace(")", string.Empty);

            var indexOfA = Array.IndexOf(Constants.INDICATOR_KEY_NAME_MAPPING.Keys.ToArray(), indicatorAcronymA);
            var indexOfB = Array.IndexOf(Constants.INDICATOR_KEY_NAME_MAPPING.Keys.ToArray(), indicatorAcronymB);

            return indexOfA < indexOfB ? -1 : 1;
        }
      
    }
}
