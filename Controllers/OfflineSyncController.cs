using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using VelastoProductionSystem.Data;
using VelastoProductionSystem.Models;

namespace VelastoProductionSystem.Controllers
{
    /// <summary>
    /// Controller publik khusus untuk sinkronisasi data ke IndexedDB.
    /// Tidak memerlukan sesi login agar tablet bisa menarik data sebelum login.
    /// </summary>
    public class OfflineSyncController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ElwpDbContext _elwpContext;

        public OfflineSyncController(ApplicationDbContext context, ElwpDbContext elwpContext)
        {
            _context = context;
            _elwpContext = elwpContext;
        }

        /// <summary>
        /// Endpoint publik: mengembalikan snapshot semua data yang dibutuhkan untuk operasi offline penuh.
        /// TIDAK memerlukan autentikasi (session) — bisa dipanggil kapan saja saat online.
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetFullSnapshot(DateTime? targetDate = null)
        {
            try
            {
                var date = (targetDate ?? DateTime.Today).Date;
                var yesterday = date.AddDays(-1);
                var tomorrow = date.AddDays(1);
                var cultureID = new System.Globalization.CultureInfo("id-ID");

                // 1. Users (untuk Login Offline)
                var users = await _elwpContext.ElwpUsers
                    .Where(u => u.IsActive && u.AreaId == 1)
                    .OrderBy(u => u.FullName)
                    .Select(u => new { u.Id, u.FullName, u.Username })
                    .AsNoTracking()
                    .ToListAsync();

                // 2. Machines (untuk Login Offline + form)
                var machines = await _elwpContext.ElwpMachines
                    .Where(m => m.IsActive && m.AreaId == 1 && m.KodeMesin != "DL01" && m.KodeMesin != "DL02")
                    .OrderBy(m => m.KodeMesin)
                    .Select(m => new { m.Id, m.KodeMesin, m.NamaMesin })
                    .AsNoTracking()
                    .ToListAsync();

                // 3. Planning (hari ini + kemarin untuk jaga-jaga shift malam)
                var elwpRows = await (
                    from p in _elwpContext.ElwpPlannings
                    join m in _elwpContext.ElwpMachines on p.MesinId equals m.Id into machineJoin
                    from m in machineJoin.DefaultIfEmpty()
                    where p.TanggalPlanning >= yesterday && p.TanggalPlanning < tomorrow
                    select new
                    {
                        p.KodeItem,
                        PartName = p.PartName ?? "#N/A",
                        p.PnSap,
                        MachineName = m != null ? m.NamaMesin : "UNKNOWN",
                        MachineCode = m != null ? m.KodeMesin : "",
                        p.TanggalPlanning,
                        p.Shift
                    }
                ).AsNoTracking().ToListAsync();

                var plannings = elwpRows.Select(x => new
                {
                    itemCode = x.KodeItem,
                    itemName = x.PartName + (string.IsNullOrEmpty(x.PnSap) ? "" : " | " + x.PnSap),
                    machineName = x.MachineName,
                    machineCode = x.MachineCode,
                    dateShift = $"{(x.TanggalPlanning?.ToString("dddd, d MMMM yyyy", cultureID) ?? "").ToUpper()} SHIFT {x.Shift}",
                    date = x.TanggalPlanning?.ToString("yyyy-MM-dd"),
                    shift = x.Shift?.ToString(),
                    isFiltered = false,
                    hasSps = true
                }).OrderBy(p => p.itemName).ToList();

                // 4. SPS Docs — load sekali, pakai untuk dua keperluan (Parameter + Dimensi)
                var spsRaw = await _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .AsNoTracking()
                    .ToListAsync();

                // 4a. SPS Parameter (struktur yang dipakai Create.cshtml)
                var spsParamList = spsRaw.Select(s => BuildSpsParamEntry(s)).ToList();

                // 4b. SPS Dimensi (struktur yang dipakai App.cshtml Dimensi)
                var spsDimList = spsRaw.Select(s => BuildSpsDimEntry(s)).ToList();

                // 5. Shift Master
                var shiftMasters = await _context.ShiftMasters
                    .AsNoTracking()
                    .Select(s => new { s.ShiftName, s.StartTime })
                    .ToListAsync();

                // 6. Summary laporan hari ini (untuk keperluan dashboard)
                var paramReports = await _context.ProductionReports
                    .Where(r => r.CreatedDate >= date && r.CreatedDate < tomorrow)
                    .Select(r => new { r.Id, r.HoseType, r.VinCode, r.Shift, r.CreatedBy, r.Status, r.CreatedDate, r.DocumentNumber, r.MachineName })
                    .AsNoTracking()
                    .ToListAsync();

                var dimReports = await _context.DimensionReports
                    .Where(r => r.CreatedDate >= date && r.CreatedDate < tomorrow)
                    .Select(r => new { r.Id, r.HoseType, r.Shift, r.CreatedBy, r.Status, r.CreatedDate, r.DocumentNumber, r.MachineName })
                    .AsNoTracking()
                    .ToListAsync();

                return Json(new
                {
                    success = true,
                    syncedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    users,
                    machines,
                    plannings,
                    spsParamList,
                    spsDimList,
                    shiftMasters,
                    summary = new { parameterReports = paramReports, dimensionReports = dimReports }
                });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // ─── Helper: SPS untuk form Parameter (Create.cshtml) ─────────────────────
        private object BuildSpsParamEntry(SpsNoDoc s)
        {
            var itemCodes = s.ItemLists?
                .Select(il => il.ItemList?.Trim().ToUpperInvariant())
                .Where(c => !string.IsNullOrEmpty(c))
                .Distinct()
                .ToList() ?? new List<string?>();

            return new
            {
                documentNumber = s.DocumentNumber,
                itemCodes,
                paramData = new
                {
                    documentNumber = s.DocumentNumber,
                    hoseType = s.HoseType,
                    customer = s.Customer,
                    nipple = s.Nipple,
                    tubeDie = s.TubeDie,
                    middleDie = s.MiddleDie,
                    outerDie = s.CoverDie,
                    innerTube = s.InnerTube,
                    middleTube = s.MiddleTube,
                    outerCover = s.OuterCover,
                    yarn = s.Yarn,
                    material = s.Material,
                    itemLists = s.ItemLists?.Select(il => new
                    {
                        il.Id,
                        il.ItemList
                    }).ToList()
                }
            };
        }

        // ─── Helper: SPS untuk form Dimensi (App.cshtml) ──────────────────────────
        private object BuildSpsDimEntry(SpsNoDoc s)
        {
            string R(decimal? asli, decimal? min, decimal? max, string? fallback = "")
            {
                if (asli.HasValue)
                {
                    string sMin = min.HasValue ? min.Value.ToString("F2") : "--";
                    string sMax = max.HasValue ? max.Value.ToString("F2") : "--";
                    return $"{sMin} | {asli.Value:F2} | {sMax}";
                }
                if (min.HasValue || max.HasValue)
                {
                    string sMin = min.HasValue ? min.Value.ToString("F2") : "--";
                    string sMax = max.HasValue ? max.Value.ToString("F2") : "--";
                    return $"{sMin} | -- | {sMax}";
                }
                string fb = fallback?.Trim() ?? "";
                var invalidFb = new[] { "", "-", "--", "---", "kosong", "null", "n/a" };
                if (invalidFb.Contains(fb.ToLower())) return "";
                if (fb.Contains("|")) return fb;
                return $"-- | {fb} | --";
            }

            return new
            {
                success = true,
                documentNumber = s.DocumentNumber,
                hoseType = s.HoseType,
                data = new
                {
                    success = true,
                    id = s.DocumentNumber,
                    documentNumber = s.DocumentNumber,
                    hoseType = s.HoseType,
                    customerName = s.Customer,
                    innerMaterial = s.InnerTube,
                    middleMaterial = s.MiddleTube,
                    outerMaterial = s.OuterCover,
                    yarnType = s.Yarn ?? s.Material,
                    layerType = !string.IsNullOrWhiteSpace(s.MiddleTube) ? "CHS 3 Layer" : "CHS 2 Layer",

                    // Die dimensions
                    innerDie = R(s.Nipple_Asli, s.Nipple_Min, s.Nipple_Max, s.Nipple),
                    nippleDie = R(s.Nipple_Asli, s.Nipple_Min, s.Nipple_Max, s.Nipple),
                    tubeDie = R(s.TubeDie_Asli, s.TubeDie_Min, s.TubeDie_Max, s.TubeDie),
                    middleDie = R(s.MiddleDie_Asli, s.MiddleDie_Min, s.MiddleDie_Max, s.MiddleDie),
                    coverDie = R(s.CoverDie_Asli, s.CoverDie_Min, s.CoverDie_Max, s.CoverDie),
                    spacerDie = R(s.SpacerDie_Asli, s.SpacerDie_Min, s.SpacerDie_Max, s.SpacerDie),
                    toleranceDie = R(s.ADistance_Asli, s.ADistance_Min, s.ADistance_Max, s.ADistance),

                    // Inner extruder
                    headTempInner = R(s.HeadTemp1_Asli, s.HeadTemp1_Min, s.HeadTemp1_Max, s.HeadTemp1),
                    cylinder1TempInner = R(s.Cylinder1_1_Asli, s.Cylinder1_1_Min, s.Cylinder1_1_Max, s.Cylinder1_1),
                    cylinder2TempInner = R(s.Cylinder2_1_Asli, s.Cylinder2_1_Min, s.Cylinder2_1_Max, s.Cylinder2_1),
                    cylinder3TempInner = R(s.Cylinder3_1_Asli, s.Cylinder3_1_Min, s.Cylinder3_1_Max, s.Cylinder3_1),
                    screwTempInner = R(s.ScrewTemp1_Asli, s.ScrewTemp1_Min, s.ScrewTemp1_Max, s.ScrewTemp1),
                    screwSpeedInner = R(s.ScrewSpeed1_Asli, s.ScrewSpeed1_Min, s.ScrewSpeed1_Max, s.ScrewSpeed1),
                    pressureInner = R(s.Pressure1_Asli, s.Pressure1_Min, s.Pressure1_Max, s.Pressure1),
                    feedRollRatioInner = R(s.FeedRollRatio1_Asli, s.FeedRollRatio1_Min, s.FeedRollRatio1_Max,
                        !string.IsNullOrEmpty(s.FeedRollRatio1) ? s.FeedRollRatio1 : s.Feed1),

                    // Outer extruder
                    headTempOuter = R(s.HeadTemp2_Asli, s.HeadTemp2_Min, s.HeadTemp2_Max, s.HeadTemp2),
                    cylinder1TempOuter = R(s.Cylinder1_2_Asli, s.Cylinder1_2_Min, s.Cylinder1_2_Max, s.Cylinder1_2),
                    cylinder2TempOuter = R(s.Cylinder2_2_Asli, s.Cylinder2_2_Min, s.Cylinder2_2_Max, s.Cylinder2_2),
                    cylinder3TempOuter = R(s.Cylinder3_2_Asli, s.Cylinder3_2_Min, s.Cylinder3_2_Max, s.Cylinder3_2),
                    screwTempOuter = R(s.ScrewTemp2_Asli, s.ScrewTemp2_Min, s.ScrewTemp2_Max, s.ScrewTemp2),
                    screwSpeedOuter = R(s.ScrewSpeed2_Asli, s.ScrewSpeed2_Min, s.ScrewSpeed2_Max, s.ScrewSpeed2),
                    pressureOuter = R(s.Pressure2_Asli, s.Pressure2_Min, s.Pressure2_Max, s.Pressure2),
                    feedRollRatioOuter = R(s.FeedRollRatio2_Asli, s.FeedRollRatio2_Min, s.FeedRollRatio2_Max,
                        !string.IsNullOrEmpty(s.FeedRollRatio2) ? s.FeedRollRatio2 : s.Feed2),

                    // Process params
                    spiralSpeed = R(s.SpiralSpeed_Asli, s.SpiralSpeed_Min, s.SpiralSpeed_Max, s.SpiralSpeed),
                    toleranceSpiralPitch = R(s.SpiralPitchSetting_Asli, s.SpiralPitchSetting_Min, s.SpiralPitchSetting_Max,
                        !string.IsNullOrEmpty(s.SpiralPitchSetting) ? s.SpiralPitchSetting : s.ToleranceSpiralPitch),
                    presetValue = R(s.PresetValue_Asli, s.PresetValue_Min, s.PresetValue_Max, s.PresetValue),
                    controlValue = R(s.ControlValue_Asli, s.ControlValue_Min, s.ControlValue_Max, s.ControlValue),
                    hoseSpeed = R(s.HoseSpeed_Asli, s.HoseSpeed_Min, s.HoseSpeed_Max, s.HoseSpeed),
                    takeupConveyorSpeed = R(s.TakeUpConveyorSpeed_Asli, s.TakeUpConveyorSpeed_Min, s.TakeUpConveyorSpeed_Max, s.TakeUpConveyorSpeed),
                    spiralPitchSetting = R(s.SpiralPitchSetting_Asli, s.SpiralPitchSetting_Min, s.SpiralPitchSetting_Max, s.SpiralPitchSetting),
                    chillerWaterTemp = R(s.ChillerWaterTemp_Asli, s.ChillerWaterTemp_Min, s.ChillerWaterTemp_Max, s.ChillerWaterTemp),
                    caterpillarGap = R(s.CaterpillarGap_Asli, s.CaterpillarGap_Min, s.CaterpillarGap_Max, s.CaterpillarGap)
                }
            };
        }
    }
}
