
using System;

namespace EpdConverter.Core.Models
{
    // TODO: Encapsulation -> private setters
    public class EpdIndicator
    {
        /* TODO: Translate comment to en
         * Unter Umständen sind die Pahsen A1-A3 separat ODER aggregiert angegeben.
         * Modeliert wird die Anweundung Jedoch immer in aggregierter Form für diese Phasen.
         * Beispiel für aggregierte EPD: Spannbeton-Fertigteildecken
         */

        public string IndicatorDescription { get; set; }

        public string Direction { get; set; }

        public string Unit { get; set; }

        /// <summary>
        /// A1 - A3
        /// </summary>
        public double? ProductionA1ToA3 { get; set; }

        /// <summary>
        /// A4
        /// </summary>
        public double? TransportA4 { get; set; }

        /// <summary>
        /// A5
        /// </summary>
        public double? BuildingProcessA5 { get; set; }

        /// <summary>
        /// B1
        /// </summary>
        public double? UsageB1 { get; set; }

        /// <summary>
        /// B2
        /// </summary>
        public double? MaintenanceB2 { get; set; }

        /// <summary>
        /// B3
        /// </summary>
        public double? RepairB3 { get; set; }

        /// <summary>
        /// B4
        /// </summary>
        public double? ReplacementB4 { get; set; }

        /// <summary>
        /// B5
        /// </summary>
        public double? ModernizationB5 { get; set; }

        /// <summary>
        /// B6
        /// </summary>
        public double? EnergyDemandB6 { get; set; }

        /// <summary>
        /// B7
        /// </summary>
        public double? WaterDemandB7 { get; set; }

        /// <summary>
        /// C1
        /// </summary>
        public double? BreakUpC1 { get; set; }

        /// <summary>
        /// C2
        /// </summary>
        public double? TransportC2 { get; set; }

        /// <summary>
        /// C3
        /// </summary>
        public double? WasteManagementC3 { get; set; }

        /// <summary>
        /// C4
        /// </summary>
        public double? WasteDisposalC4 { get; set; }

        /// <summary>
        /// D
        /// </summary>
        public double? ReuseAndRecoveryD { get; set; }
    }
}
