namespace VelastoProductionSystem.Models
{
    public class ProductionMaterialLotDto
    {
        public string? LayerType { get; set; }
        public string? MaterialName { get; set; }
        public string? MaterialActual { get; set; }
        public string? LotNumber { get; set; }
        public decimal? SGValue { get; set; }
    }

    public class DimensionReportAppData
    {
        public int ReportId { get; set; }
        public string? DocumentNumber { get; set; }
        public string? ItemCode { get; set; }
        public string? HoseType { get; set; }
        public string? MachineName { get; set; }
        public string? DimensionDisplay { get; set; }
        public string? CustomerName { get; set; }
        public string? Yarn { get; set; }
        public string? VinCode { get; set; }
        public string? StandardLength { get; set; }
        public string? StandardShrink { get; set; }
        public string? ActualLength { get; set; }
        public int? QtyTarget { get; set; }
        public int? QtyOk { get; set; }
        public int? NgDimension { get; set; }
        public int? NgVisual { get; set; }
        public string? Remark { get; set; }
        public string? ByPass { get; set; }
        public string? Shift { get; set; }
        public int RevisionNumber { get; set; }
        public string? CreatedBy { get; set; }
        public string? SpsId { get; set; }
        
        // Dimension readings array
        public List<DimensionReadingData> DimensionReadings { get; set; } = new List<DimensionReadingData>();
        
        // Summaries array
        public List<DimensionSummaryData> ProductionDataSummaries { get; set; } = new List<DimensionSummaryData>();
    }

    public class DimensionSummaryData
    {
        public string? PartNumber { get; set; }
        public string? VinCode { get; set; }
        public string? StandardLength { get; set; }
        public string? StandardShrink { get; set; }
        public string? ActualLength { get; set; }
        public int? QtyTarget { get; set; }
        public int? QtyOk { get; set; }
        public int? NgDimension { get; set; }
        public int? NgVisual { get; set; }
    }

    public class DimensionReadingData
    {
        public string? PointName { get; set; }
        public string? TimeSection { get; set; }
        public string? Frequency { get; set; }
        public string? StandardDimension { get; set; }
        public string? Initial { get; set; }
        public string? Reading1 { get; set; }
        public string? Reading2 { get; set; }
        public string? Reading3 { get; set; }
        public string? Reading4 { get; set; }
        public string? Reading5 { get; set; }
        public string? Reading6 { get; set; }
        public string? Status { get; set; }
        public decimal? ScaleValue { get; set; }
    }

    public class ExtruderProducePayload
    {
        public int Id { get; set; }
        public string? HoseType { get; set; }
        public string? Dimension { get; set; }
        public string? Yarn { get; set; }
        public string? MaterialInner { get; set; }
        public string? MaterialInnerLotNo { get; set; }
        public decimal? MaterialInnerSG { get; set; }
        public string? MaterialMiddle { get; set; }
        public string? MaterialMiddleLotNo { get; set; }
        public decimal? MaterialMiddleSG { get; set; }
        public string? MaterialOuter { get; set; }
        public string? MaterialOuterLotNo { get; set; }
        public decimal? MaterialOuterSG { get; set; }
        public DateTime? DandoriStart { get; set; }
        public DateTime? DandoriEnd { get; set; }
        public DateTime? ProductionStart { get; set; }
        public bool SPVCheck { get; set; }
    }

    public class ProductionReportSaveDto
    {
        public int Id { get; set; }
        public string? DocumentNumber { get; set; }
        public int RevisionNumber { get; set; }
        public string? ProductionDate { get; set; }
        public string? Shift { get; set; }
        public string? CustomerName { get; set; }
        public string? HoseType { get; set; }
        public string? InnerMaterial { get; set; }
        public string? InnerMaterialActual { get; set; }
        public string? InnerMaterialLotNo { get; set; }
        public string? OuterMaterial { get; set; }
        public string? OuterMaterialActual { get; set; }
        public string? OuterMaterialLotNo { get; set; }
        public string? MiddleMaterial { get; set; }
        public string? MiddleMaterialActual { get; set; }
        public string? MiddleMaterialLotNo { get; set; }
        public string? Yarn { get; set; }
        public string? YarnActual { get; set; }
        public string? YarnLotNo { get; set; }
        public string? DAI_Awal { get; set; }
        public string? DAI_Akhir { get; set; }
        public string? DAC_Awal { get; set; }
        public string? DAC_Akhir { get; set; }
        public string? DRAI_Awal { get; set; }
        public string? DRAI_Akhir { get; set; }
        public string? DRAC_Awal { get; set; }
        public string? DRAC_Akhir { get; set; }
        public string? CreatedBy { get; set; }
        
        // --- NEW FIELDS ---
        public string? SpsId { get; set; }
        public string? ItemCode { get; set; }
        
        // Dies
        public bool NippleDieOK { get; set; }
        public string? NippleDieInitial { get; set; }
        public string? NippleDieFinal { get; set; }
        public bool TubeDieOK { get; set; }
        public string? TubeDieInitial { get; set; }
        public string? TubeDieFinal { get; set; }
        public bool MiddleDieOK { get; set; }
        public string? MiddleDieInitial { get; set; }
        public string? MiddleDieFinal { get; set; }
        public bool CoverDieOK { get; set; }
        public string? CoverDieInitial { get; set; }
        public string? CoverDieFinal { get; set; }
        public bool SpacerDieOK { get; set; }
        public string? SpacerDieInitial { get; set; }
        public string? SpacerDieFinal { get; set; }
        public bool ToleranceDieOK { get; set; }
        public string? ToleranceInitial { get; set; }
        public string? ToleranceFinal { get; set; }
        public bool MeshInner10Before { get; set; }
        public bool MeshInner40Before { get; set; }
        public bool MeshOuter10Before { get; set; }
        public bool MeshOuter40Before { get; set; }
        public string? MeshInnerCheck { get; set; }
        public string? MeshOuterCheck { get; set; }
        public string? EmbossMarkContent { get; set; }
        public string? EmbossMarkDate { get; set; }
        public string? QcCond { get; set; }
        public string? QcSurf { get; set; }
        public string? QcRes { get; set; }
        public string? AlarmFunctionCheck { get; set; }

        // Item List Lengths
        public decimal? LengthShrinkMin { get; set; }
        public decimal? LengthShrinkStd { get; set; }
        public decimal? LengthShrinkMax { get; set; }
        public decimal? LengthCuttingMin { get; set; }
        public decimal? LengthCuttingStd { get; set; }
        public decimal? LengthCuttingMax { get; set; }
        public decimal? ActualLengthShrink { get; set; }
        public decimal? ActualLengthCutting { get; set; }

        public List<ProductionReadingSaveDto> Readings { get; set; } = new List<ProductionReadingSaveDto>();
        public List<ProductionMaterialLotDto> MaterialLots { get; set; } = new List<ProductionMaterialLotDto>();
    }

    public class ProductionReadingSaveDto
    {
        public string? ReadingTime { get; set; }
        public string? RecordedBy { get; set; }
        public string? HeadTempInner { get; set; }
        public string? Cylinder1TempInner { get; set; }
        public string? Cylinder2TempInner { get; set; }
        public string? Cylinder3TempInner { get; set; }
        public string? ScrewTempInner { get; set; }
        public string? ScrewSpeedInner { get; set; }
        public string? FeedRollRatioInner { get; set; }
        public string? PressureInner { get; set; }
        public string? AmMeterInner { get; set; }

        // Middle (Tube 2)
        public string? HeadTempMiddle { get; set; }
        public string? Cylinder1TempMiddle { get; set; }
        public string? Cylinder2TempMiddle { get; set; }
        public string? Cylinder3TempMiddle { get; set; }
        public string? ScrewTempMiddle { get; set; }
        public string? ScrewSpeedMiddle { get; set; }
        public string? FeedRollRatioMiddle { get; set; }
        public string? PressureMiddle { get; set; }
        public string? AmMeterMiddle { get; set; }

        // Outer
        public string? HeadTempOuter { get; set; }
        public string? Cylinder1TempOuter { get; set; }
        public string? Cylinder2TempOuter { get; set; }
        public string? Cylinder3TempOuter { get; set; }
        public string? ScrewTempOuter { get; set; }
        public string? ScrewSpeedOuter { get; set; }
        public string? FeedRollRatioOuter { get; set; }
        public string? PressureOuter { get; set; }
        public string? AmMeterOuter { get; set; }

        // Spiral
        public string? SpiralSpeed { get; set; }
        public string? SpiralPitchSetting { get; set; }
        public string? SpiralPitchDisplay { get; set; }

        // Infra & Utility
        public string? PresetValue { get; set; }
        public string? ControlValue { get; set; }
        public string? OdSensor { get; set; }
        public string? HoseSpeed { get; set; }
        public string? TakeupConveyorSpeed { get; set; }
        public string? CoolConveyorSpeed { get; set; }
        public string? CoolConveyorSpeed2 { get; set; }
        public string? ConveyorRatio { get; set; }
        public string? UnsmoothSurface { get; set; }
        public string? ChillerWaterTemp { get; set; }
        public string? CaterpillarGap { get; set; }
        public string? CuttingSpeed { get; set; }
    }
}
