using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;

namespace VelastoProductionSystem.Models
{
    public class DimensionSummary
    {
        public int Id { get; set; }
        public int DimensionReportId { get; set; }
        
        public string? PartNumber { get; set; }
        public string? VinCode { get; set; }
        public string? StandardLength { get; set; }
        public string? StandardShrink { get; set; }
        public string? ActualLength { get; set; }
        public int QtyTarget { get; set; }
        public int QtyOk { get; set; }
        public int NgDimension { get; set; }
        public int NgVisual { get; set; }

        [JsonIgnore]
        public DimensionReport? DimensionReport { get; set; }
    }
}
