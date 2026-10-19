using System.ComponentModel.DataAnnotations;

namespace VelastoProductionSystem.Models
{
    public class ProductionReading
    {
        public int Id { get; set; }

        [Required]
        public int ProductionReportId { get; set; }
        public ProductionReport? ProductionReport { get; set; }

        [Required]
        [Display(Name = "Waktu Pengukuran")]
        public DateTime ReadingTime { get; set; }

        [Display(Name = "Interval (menit)")]
        public int IntervalMinutes { get; set; }

        // Temperature Readings (Inner Head)
        [Display(Name = "Head Temp Inner (°C)")]
        public int? HeadTempInner { get; set; }

        [Display(Name = "Cylinder 1 Inner (°C)")]
        public int? Cylinder1TempInner { get; set; }

        [Display(Name = "Cylinder 2 Inner (°C)")]
        public int? Cylinder2TempInner { get; set; }

        [Display(Name = "Cylinder 3 Inner (°C)")]
        public int? Cylinder3TempInner { get; set; }

        [Display(Name = "Screw Temp Inner (°C)")]
        public int? ScrewTempInner { get; set; }

        // Speed Settings (Inner)
        [Display(Name = "Screw Speed Inner (rpm)")]
        public decimal? ScrewSpeedInner { get; set; }

        [Display(Name = "Feed Roll Ratio Inner (%)")]
        public int? FeedRollRatioInner { get; set; }

        [Display(Name = "Pressure Inner (MPa)")]
        public decimal? PressureInner { get; set; }

        // Am meter (Inner)
        [Display(Name = "Am meter Inner (A)")]
        public decimal? AmMeterInner { get; set; }

        // Temperature Readings (Middle Head - Tube 2)
        [Display(Name = "Head Temp Middle (°C)")]
        public int? HeadTempMiddle { get; set; }

        [Display(Name = "Cylinder 1 Middle (°C)")]
        public int? Cylinder1TempMiddle { get; set; }

        [Display(Name = "Cylinder 2 Middle (°C)")]
        public int? Cylinder2TempMiddle { get; set; }

        [Display(Name = "Cylinder 3 Middle (°C)")]
        public int? Cylinder3TempMiddle { get; set; }

        [Display(Name = "Screw Temp Middle (°C)")]
        public int? ScrewTempMiddle { get; set; }

        // Speed & Pressure Settings (Middle - Tube 2)
        [Display(Name = "Screw Speed Middle (rpm)")]
        public decimal? ScrewSpeedMiddle { get; set; }

        [Display(Name = "Feed Roll Ratio Middle (%)")]
        public int? FeedRollRatioMiddle { get; set; }

        [Display(Name = "Pressure Middle (MPa)")]
        public decimal? PressureMiddle { get; set; }

        [Display(Name = "Am meter Middle (A)")]
        public decimal? AmMeterMiddle { get; set; }

        // Temperature Readings (Outer Head)
        [Display(Name = "Head Temp Outer (°C)")]
        public int? HeadTempOuter { get; set; }

        [Display(Name = "Cylinder 1 Outer (°C)")]
        public int? Cylinder1TempOuter { get; set; }

        [Display(Name = "Cylinder 2 Outer (°C)")]
        public int? Cylinder2TempOuter { get; set; }

        [Display(Name = "Cylinder 3 Outer (°C)")]
        public int? Cylinder3TempOuter { get; set; }

        [Display(Name = "Screw Temp Outer (°C)")]
        public int? ScrewTempOuter { get; set; }

        // Speed Settings (Outer)
        [Display(Name = "Screw Speed Outer (rpm)")]
        public decimal? ScrewSpeedOuter { get; set; }

        [Display(Name = "Feed Roll Ratio Outer (%)")]
        public int? FeedRollRatioOuter { get; set; }

        [Display(Name = "Pressure Outer (MPa)")]
        public decimal? PressureOuter { get; set; }

        [Display(Name = "Am meter Outer (A)")]
        public decimal? AmMeterOuter { get; set; }

        // Spiral
        [Display(Name = "Spiral Speed (rpm)")]
        public decimal? SpiralSpeed { get; set; }

        [Display(Name = "Spiral Pitch Setting")]
        public decimal? SpiralPitchSetting { get; set; }

        [Display(Name = "Spiral Pitch Display")]
        public decimal? SpiralPitchDisplay { get; set; }

        [Display(Name = "Preset Temp (°C)")]
        public decimal? PresetTemp { get; set; }

        [Display(Name = "Preset Value (Ø mm)")]
        public decimal? PresetValue { get; set; }

        // Control Values
        [Display(Name = "Control Value (Ø mm)")]
        public decimal? ControlValue { get; set; }

        [Display(Name = "OD Sensor (Ø mm)")]
        public decimal? OdSensor { get; set; }

        [Display(Name = "Hose Speed (m/min)")]
        public decimal? HoseSpeed { get; set; }

        [Display(Name = "Take-up Conveyor Speed (m/min)")]
        public decimal? TakeupConveyorSpeed { get; set; }

        [Display(Name = "Cool Conveyor Speed (m/min)")]
        public decimal? CoolConveyorSpeed { get; set; }

        [Display(Name = "Cool Conveyor 2 Speed (m/min)")]
        public decimal? CoolConveyorSpeed2 { get; set; }

        [Display(Name = "Conveyor Ratio")]
        public string? ConveyorRatio { get; set; }

        [Display(Name = "Un-smooth Surface")]
        public decimal? UnsmoothSurface { get; set; }

        [Display(Name = "Chiller Water Temp (°C)")]
        public decimal? ChillerWaterTemp { get; set; }

        [Display(Name = "Caterpillar Gap (mm)")]
        public decimal? CaterpillarGap { get; set; }

        [Display(Name = "Cutting Speed (m/min)")]
        public decimal? CuttingSpeed { get; set; }

        // DIMENSION REPORT FIELDS (Gambar Dimension)
        [Display(Name = "Inner Diameter (mm)")]
        public decimal? InnerDiameter { get; set; }

        [Display(Name = "Inner Thickness X (mm)")]
        public decimal? InnerThicknessX { get; set; }

        [Display(Name = "Inner Thickness Y (mm)")]
        public decimal? InnerThicknessY { get; set; }

        [Display(Name = "Total Thickness X (mm)")]
        public decimal? TotalThicknessX { get; set; }

        [Display(Name = "Total Thickness Y (mm)")]
        public decimal? TotalThicknessY { get; set; }

        [Display(Name = "Spiral Pitch (mm)")]
        public decimal? SpiralPitch { get; set; }

        [Display(Name = "Visual Check")]
        public string? VisualCheck { get; set; }

        // Notes
        [Display(Name = "Catatan")]
        public string? Notes { get; set; }

        [Display(Name = "Dicatat Oleh")]
        public string RecordedBy { get; set; } = string.Empty;
    }
}
