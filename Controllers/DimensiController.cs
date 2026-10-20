using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using VelastoProductionSystem.Data;
using VelastoProductionSystem.Models;
using System.Text.Json;

using VelastoProductionSystem.Hubs;
using VelastoProductionSystem.Helpers;
using Microsoft.AspNetCore.SignalR;

namespace VelastoProductionSystem.Controllers
{
    public class DimensiController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ElwpDbContext _elwpContext;
        private readonly IHubContext<DashboardHub> _hubContext;

        public DimensiController(ApplicationDbContext context, ElwpDbContext elwpContext, IHubContext<DashboardHub> hubContext)
        {
            _context = context;
            _elwpContext = elwpContext;
            _hubContext = hubContext;
        }

        [Microsoft.AspNetCore.Authorization.AllowAnonymous]
        public async Task<IActionResult> DebugSps()
        {
            var sps = await _context.SpsNoDocs.FirstOrDefaultAsync(s => s.DocumentNumber == "VI-SOP-PROD-132");
            return Json(sps);
        }

        // GET: /Dimensi/Index
        public IActionResult Index()
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("UserName")))
            {
                return RedirectToAction("Login", "Account");
            }

            // Ambil Kode Item dari ELWP langsung (hari ini), agar selalu up-to-date
            var today = DateTime.Today;
            var sessionMachine = HttpContext.Session.GetString("MachineName");
            var isAdmin = HttpContext.Session.GetString("IsAdmin") == "true";

            var elwpQuery = from p in _elwpContext.ElwpPlannings
                            join m in _elwpContext.ElwpMachines on p.MesinId equals m.Id into machineJoin
                            from m in machineJoin.DefaultIfEmpty()
                            where !string.IsNullOrEmpty(p.KodeItem)
                               && p.TanggalPlanning >= today
                               && p.TanggalPlanning < today.AddDays(1)
                            select new { p.KodeItem, MachineName = m != null ? m.NamaMesin : "", MachineCode = m != null ? m.KodeMesin : "" };

            var elwpRows = elwpQuery.ToList();

            if (!isAdmin && !string.IsNullOrEmpty(sessionMachine))
            {
                var normalizedFilter = new string(sessionMachine.Where(char.IsLetterOrDigit).ToArray()).ToUpper();
                elwpRows = elwpRows.Where(x =>
                    (x.MachineName != null && new string(x.MachineName.Where(char.IsLetterOrDigit).ToArray()).ToUpper().Contains(normalizedFilter)) ||
                    (x.MachineCode != null && new string(x.MachineCode.Where(char.IsLetterOrDigit).ToArray()).ToUpper().Contains(normalizedFilter))
                ).ToList();
            }

            ViewBag.ItemCodes = elwpRows.Select(x => x.KodeItem)
                .Distinct()
                .OrderBy(p => p)
                .ToList();

            var report = new DimensionReport
            {
                DocumentNumber = "",
                ProductionDate = DateTime.Now,
                Shift = GetCurrentShift(),
                Status = "DRAFT",
                ItemCode = "",
                CreatedBy = HttpContext.Session.GetString("UserName") ?? "QC Operator",
                MachineName = HttpContext.Session.GetString("MachineName")
            };
            return View("App", report);
        }

        private string GetCurrentShift() =>
            ShiftHelper.GetCurrentShift(_context.ShiftMasters.ToList());

        [HttpGet]
        public async Task<JsonResult> GetSpsStandard(string hoseType, string? machineCode)
        {
            if (string.IsNullOrEmpty(hoseType)) return Json(new { success = false, message = "Input empty" });

            var sanitized = hoseType.Trim().ToUpper();
            var sanitizedMachine = string.IsNullOrEmpty(machineCode) ? null : machineCode.Trim().ToUpper();
            var sanitizedItemLookup = sanitized.Replace("-", "").Replace(" ", "");

            string ResolveSpiralPitchTarget(SpsNoDoc src)
            {
                var target = src.SpiralPitchSetting_Asli?.ToString("F2") ?? src.SpiralPitchDisplay_Asli?.ToString("F2") ?? src.SpiralPitchSetting;
                if (string.IsNullOrWhiteSpace(target) || target == "0") target = src.SpiralPitchDisplay;
                if (string.IsNullOrWhiteSpace(target) || target == "0") target = src.ToleranceSpiralPitch;
                return target ?? "";
            }

            string ResolveSpiralPitchTolerance(SpsNoDoc src)
            {
                var tol = src.ToleranceSpiralPitch_Asli?.ToString("F2") ?? src.ToleranceSpiralPitch;
                if (string.IsNullOrWhiteSpace(tol)) tol = "5";
                return tol;
            }

            string ResolveSpiralPitchDisplay(SpsNoDoc src)
            {
                var disp = src.SpiralPitchDisplay_Asli?.ToString("F2") ?? src.SpiralPitchDisplay;
                if (string.IsNullOrWhiteSpace(disp) || disp == "0") disp = src.SpiralPitchSetting_Asli?.ToString("F2") ?? src.SpiralPitchSetting;
                return disp ?? "";
            }

            // 1. Priority 1: Exact match by SpsItemList.ItemList (kode item dari planning)
            var spsByItemList = await _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .OrderByDescending(s => s.DocumentNumber)
                .FirstOrDefaultAsync(s =>
                    s.ItemLists.Any(il =>
                        il.ItemList != null &&
                        il.ItemList.Replace("-", "").Replace(" ", "").ToUpper() == sanitizedItemLookup));

            var sps = spsByItemList;

            // 1b. Fallback: partial item list match (legacy behavior) if exact match is not found
            if (sps == null)
            {
                sps = await _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .OrderByDescending(s => s.DocumentNumber)
                    .FirstOrDefaultAsync(s =>
                        s.ItemLists.Any(il =>
                            il.ItemList != null &&
                            il.ItemList.Replace("-", "").Replace(" ", "").ToUpper().Contains(sanitizedItemLookup)));
            }

            // 2. Try SpsMasters (Master SPS Edit) FIRST - prioritized as per mentor direction
            if (sps == null)
            {
                sps = await _context.SpsNoDocs
                    .FirstOrDefaultAsync(s => 
                        (s.DocumentNumber != null && s.DocumentNumber.Trim().ToUpper() == sanitized) || 
                        (s.HoseType != null && s.HoseType.Trim().ToUpper() == sanitized) ||
                        (s.DocumentNumber != null && s.DocumentNumber.ToUpper().Contains(sanitized)) ||
                        (s.HoseType != null && s.HoseType.ToUpper().Contains(sanitized)));
            }

            if (sps != null)
            {
                // Robust fallback for Inner Diameter
                var finalInner = sps.ToleranceInner_Asli?.ToString("F2") ?? sps.InnerTarget;
                if (string.IsNullOrEmpty(finalInner) || finalInner == "0") finalInner = sps.Dimensi;

                // Robust fallback for Inner Thickness
                var finalThick = sps.TebalInner_Asli?.ToString("F2") ?? sps.ThickTarget;
                if (string.IsNullOrEmpty(finalThick) || finalThick == "0") finalThick = sps.TebalInner;

                // Robust fallback for Total Thickness
                var finalTotal = sps.TebalTotal_Asli?.ToString("F2") ?? sps.TotalTarget;
                if (string.IsNullOrEmpty(finalTotal) || finalTotal == "0") finalTotal = sps.TebalTotal;

                var finalSpiralPitch = ResolveSpiralPitchTarget(sps);
                var finalSpiralTolerance = ResolveSpiralPitchTolerance(sps);
                var finalSpiralDisplay = ResolveSpiralPitchDisplay(sps);

                // Populate Length Shrink from SpsItemList if available
                float? shrinkMin = null, shrinkStd = null, shrinkMax = null;
                float? cuttingMin = null, cuttingStd = null, cuttingMax = null;
                var itemListsObj = new List<object>();

                if (sps.ItemLists != null && sps.ItemLists.Any())
                {
                    var matchedList = sps.ItemLists.FirstOrDefault(il => il.ItemList != null && il.ItemList.Replace("-", "").Replace(" ", "").ToUpper() == sanitizedItemLookup);
                    if (matchedList == null)
                    {
                        matchedList = sps.ItemLists.FirstOrDefault(il => il.ItemList != null && il.ItemList.Replace("-", "").Replace(" ", "").ToUpper().Contains(sanitizedItemLookup));
                    }
                    if (matchedList != null)
                    {
                        shrinkMin = matchedList.LengthShrinkMin;
                        shrinkStd = matchedList.LengthShrinkStd;
                        shrinkMax = matchedList.LengthShrinkMax;
                        cuttingMin = matchedList.LengthCuttingMin;
                        cuttingStd = matchedList.LengthCuttingStd;
                        cuttingMax = matchedList.LengthCuttingMax;
                    }

                    foreach (var il in sps.ItemLists)
                    {
                        itemListsObj.Add(new {
                            itemList = il.ItemList,
                            cuttingMin = il.LengthCuttingMin,
                            cuttingStd = il.LengthCuttingStd,
                            cuttingMax = il.LengthCuttingMax,
                            shrinkMin = il.LengthShrinkMin,
                            shrinkStd = il.LengthShrinkStd,
                            shrinkMax = il.LengthShrinkMax
                        });
                    }
                }

                return Json(new { 
                    success = true,
                    id = sps.DocumentNumber,
                    data = new {
                        hoseType = sps.HoseType,
                        dimensi = sps.Dimensi ?? finalInner, // Send the full dimension text!
                        innerTarget = finalInner,            // For quality matrix cards
                        innerTube = sps.InnerTube,
                        outerCover = sps.OuterCover,
                        toleranceInner = sps.ToleranceInner,
                        toleranceOuter = sps.ToleranceOuter,
                        tebalInner = finalThick,             // Unified Inner Thickness target
                        tebalTotal = finalTotal,             // Unified Total Thickness target
                        yarn = sps.Yarn ?? "GENERAL",
                        customer = sps.Customer,
                        docNo = sps.DocumentNumber,
                        revNo = sps.RevisionNumber,
                        itemList = hoseType,
                        toleranceSpiralPitch = finalSpiralTolerance,
                        spiralPitchSetting = finalSpiralPitch,
                        spiralPitchDisplay = finalSpiralDisplay,
                        innerTol = sps.InnerTol,
                        thickTol = sps.ThickTol,
                        totalTol = sps.TotalTol,
                        bypass = "2.0 mm",
                        itemLists = itemListsObj,
                        shrinkMin = shrinkMin,
                        shrinkStd = shrinkStd,
                        shrinkMax = shrinkMax,
                        cuttingMin = cuttingMin,
                        cuttingStd = cuttingStd,
                        cuttingMax = cuttingMax
                    }
                });
            }

            // 2. Fallback to MasterlistSpsDoubleLayers (LEGACY)
            var query = _context.SpsNoDocs.AsQueryable();
            
            if (!string.IsNullOrEmpty(sanitizedMachine))
            {
                query = query.Where(m => m.MachineCode != null && m.MachineCode.ToUpper() == sanitizedMachine);
            }

            IQueryable<SpsNoDoc> ApplyMasterlistMatch(IQueryable<SpsNoDoc> q)
            {
                return q
                    .Where(m =>
                        (m.HoseType != null && m.HoseType.Trim().ToUpper() == sanitized) ||
                        (m.DocumentNumber != null && m.DocumentNumber.Trim().ToUpper() == sanitized) ||
                        (m.DocumentNumber != null && m.DocumentNumber.ToUpper().Contains(sanitized)) ||
                        (m.HoseType != null && m.HoseType.ToUpper().Contains(sanitized))
                    )
                    .OrderByDescending(m => m.DocumentNumber != null && m.DocumentNumber.Trim().ToUpper() == sanitized)
                    .ThenByDescending(m => m.HoseType != null && m.HoseType.Trim().ToUpper() == sanitized)
                    .ThenByDescending(m => m.DocumentNumber != null && m.DocumentNumber.ToUpper().Contains(sanitized))
                    .ThenByDescending(m => m.HoseType != null && m.HoseType.ToUpper().Contains(sanitized))
                    .ThenByDescending(m => m.DocumentNumber);
            }

            var standard = await ApplyMasterlistMatch(query).FirstOrDefaultAsync();
            
            if (standard == null && !string.IsNullOrEmpty(sanitizedMachine))
            {
                standard = await ApplyMasterlistMatch(_context.SpsNoDocs).FirstOrDefaultAsync();
            }

            if (standard != null)
            {
                var finalSpiralPitch = ResolveSpiralPitchTarget(standard);
                var finalSpiralTolerance = ResolveSpiralPitchTolerance(standard);
                var finalSpiralDisplay = ResolveSpiralPitchDisplay(standard);

                return Json(new { 
                    success = true, 
                    data = new {
                        hoseType = standard.HoseType,
                        dimensi = standard.Dimensi,
                        cuttingLength = standard.CuttingSpeed, 
                        innerTube = standard.InnerTube,
                        outerCover = standard.OuterCover,
                        toleranceInner = standard.ToleranceInner,
                        toleranceOuter = standard.ToleranceOuter,
                        tebalInner = standard.TebalInner,
                        tebalInnerMiddle = standard.TebalInnerMiddle,
                        tebalTotal = standard.TebalTotal,
                        yarn = standard.Material ?? "ARAMID", 
                        customer = standard.Customer,
                        docNo = standard.DocumentNumber,
                        revNo = standard.RevisionNumber,
                        itemList = hoseType,
                        toleranceSpiralPitch = finalSpiralTolerance,
                        spiralPitchSetting = finalSpiralPitch,
                        spiralPitchDisplay = finalSpiralDisplay,
                        innerTol = standard.InnerTol,
                        thickTol = standard.ThickTol,
                        totalTol = standard.TotalTol,
                        bypass = "2.0 mm" 
                    }
                });
            }

            return Json(new { success = false, message = "Standard not found: " + hoseType });
        }

        [HttpGet]
        public async Task<IActionResult> GetAllOfflineMasterData()
        {
            try
            {
                var spsRaw = await _context.SpsNoDocs
                    .AsNoTracking()
                    .ToListAsync();
                
                var spsList = new List<object>();
                foreach (var s in spsRaw)
                {
                    spsList.Add(BuildDimensiSpsResponse(s));
                }
                
                return Json(new { success = true, spsList });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        public IActionResult App()
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("UserName")))
            {
                return RedirectToAction("Login", "Account");
            }
            return View(new DimensionReport { 
                Shift = GetCurrentShift(),
                Measurements = new List<DimensionMeasurement>() 
            });
        }

        public async Task<IActionResult> History(string? machineFilter = null)
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("UserName")))
            {
                return RedirectToAction("Login", "Account");
            }

            var query = _context.DimensionReports.AsQueryable();

            if (!string.IsNullOrWhiteSpace(machineFilter))
            {
                query = query.Where(r => r.MachineName == machineFilter);
                ViewBag.MachineFilter = machineFilter;
            }
            else
            {
                ViewBag.MachineFilter = "";
            }

            ViewBag.MachineList = await _context.DimensionReports
                .Where(r => !string.IsNullOrEmpty(r.MachineName))
                .Select(r => r.MachineName!)
                .Distinct()
                .ToListAsync();

            var reports = await query
                .OrderByDescending(d => d.CreatedDate)
                .ToListAsync();

            // Fetch SPS Map for display consistency (fixes incomplete planning/hose names)
            // Build lookup keyed by BOTH DocumentNumber AND ItemList values for numeric codes (e.g. "2950" -> "NA2950")
            var itemCodes = reports.Where(r => !string.IsNullOrEmpty(r.ItemCode)).Select(r => r.ItemCode!).Distinct().ToList();
            var allSps = await _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .ToListAsync();

            // Build dictionary: key = itemCode stored in report, value = matching SpsNoDoc
            var spsMap = new Dictionary<string, SpsNoDoc>(StringComparer.OrdinalIgnoreCase);
            foreach (var code in itemCodes)
            {
                if (spsMap.ContainsKey(code)) continue;
                // Try exact DocumentNumber match first
                var match = allSps.FirstOrDefault(s => s.DocumentNumber != null && s.DocumentNumber.Equals(code, StringComparison.OrdinalIgnoreCase));
                // Then try ItemList match (handles numeric code like "2950" matching ItemList "NA2950")
                if (match == null)
                    match = allSps.FirstOrDefault(s => s.ItemLists != null && s.ItemLists.Any(il =>
                        il.ItemList != null && (
                            il.ItemList.Equals(code, StringComparison.OrdinalIgnoreCase) ||
                            il.ItemList.TrimStart('N','A','n','a','B','b').Equals(code, StringComparison.OrdinalIgnoreCase)
                        )));
                if (match != null) spsMap[code] = match;
            }
            
            ViewBag.SpsMap = spsMap;

            return View(reports);
        }

        [HttpGet]
        public async Task<IActionResult> ExportHistoryExcel(string? searchType = null, string? query = null, string? startDate = null, string? endDate = null)
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("UserName")))
            {
                return RedirectToAction("Login", "Account");
            }

            var q = _context.DimensionReports
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(startDate) && DateTime.TryParse(startDate, out var sDate))
            {
                q = q.Where(r => r.CreatedDate >= sDate.Date);
            }

            if (!string.IsNullOrWhiteSpace(endDate) && DateTime.TryParse(endDate, out var eDate))
            {
                q = q.Where(r => r.CreatedDate <= eDate.Date.AddDays(1).AddTicks(-1));
            }

            if (!string.IsNullOrWhiteSpace(query))
            {
                var s = query.Trim().ToUpper();
                var mode = (searchType ?? "").Trim().ToLower();

                if (mode == "planning")
                {
                    q = q.Where(r =>
                        (r.ItemCode != null && r.ItemCode.ToUpper().Contains(s))
                    );
                }
                else if (mode == "operator")
                {
                    q = q.Where(r => r.CreatedBy != null && r.CreatedBy.ToUpper().Contains(s));
                }
                else if (mode == "docvin")
                {
                    q = q.Where(r =>
                        (r.DocumentNumber != null && r.DocumentNumber.ToUpper().Contains(s)) ||
                        (r.VinCode != null && r.VinCode.ToUpper().Contains(s))
                    );
                }
                else
                {
                    q = q.Where(r =>
                        (r.DocumentNumber != null && r.DocumentNumber.ToUpper().Contains(s)) ||
                        (r.VinCode != null && r.VinCode.ToUpper().Contains(s)) ||
                        (r.ItemCode != null && r.ItemCode.ToUpper().Contains(s)) ||
                        (r.CreatedBy != null && r.CreatedBy.ToUpper().Contains(s)) ||
                        (r.HoseType != null && r.HoseType.ToUpper().Contains(s))
                    );
                }
            }

            var rows = await q
                .OrderByDescending(r => r.CreatedDate)
                .ToListAsync();

            using var package = new OfficeOpenXml.ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Dimension History");

            ws.Cells[1, 1].Value = "DIMENSION HISTORY EXPORT";
            ws.Cells[1, 1, 1, 11].Merge = true;
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Font.Size = 14;
            ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(15, 23, 42));
            ws.Cells[1, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);

            ws.Cells[2, 1].Value = $"Exported at: {DateTime.Now:dd MMM yyyy HH:mm} | Total rows: {rows.Count}";
            ws.Cells[2, 1, 2, 11].Merge = true;
            ws.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells[2, 1].Style.Font.Color.SetColor(System.Drawing.Color.Gray);

            var headers = new[]
            {
                "Document", "Date", "Planning", "Input By", "Machine", "Product Info", "Passed", "Failed", "VIN", "Shift", "Customer"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                var cell = ws.Cells[4, i + 1];
                cell.Value = headers[i];
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(30, 64, 175));
                cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.FromArgb(59, 130, 246));
                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }

            // Fetch SPS Map for Excel - same dual-key lookup as History view
            var itemCodes = rows.Where(r => !string.IsNullOrEmpty(r.ItemCode)).Select(r => r.ItemCode!).Distinct().ToList();
            var allSps = await _context.SpsNoDocs.Include(s => s.ItemLists).ToListAsync();
            var spsMap = new Dictionary<string, SpsNoDoc>(StringComparer.OrdinalIgnoreCase);
            foreach (var code in itemCodes)
            {
                if (spsMap.ContainsKey(code)) continue;
                var match = allSps.FirstOrDefault(s => s.DocumentNumber != null && s.DocumentNumber.Equals(code, StringComparison.OrdinalIgnoreCase));
                if (match == null)
                    match = allSps.FirstOrDefault(s => s.ItemLists != null && s.ItemLists.Any(il =>
                        il.ItemList != null && (
                            il.ItemList.Equals(code, StringComparison.OrdinalIgnoreCase) ||
                            il.ItemList.TrimStart('N','A','n','a','B','b').Equals(code, StringComparison.OrdinalIgnoreCase)
                        )));
                if (match != null) spsMap[code] = match;
            }

            int row = 5;
            foreach (var item in rows)
            {
                var sps = (!string.IsNullOrEmpty(item.ItemCode) && spsMap.TryGetValue(item.ItemCode, out var s)) ? s : null;
                var rawPlanning = (item.ItemCode ?? string.Empty).Trim();
                var rawHose = (item.HoseType ?? string.Empty).Trim();

                // Robust Display Selection
                var spsItem = sps?.ItemLists?.FirstOrDefault(il => 
                    il.ItemList != null && (
                        il.ItemList.Equals(rawPlanning, StringComparison.OrdinalIgnoreCase) ||
                        il.ItemList.TrimStart('N','A','n','a','B','b').Equals(rawPlanning, StringComparison.OrdinalIgnoreCase)
                    ))?.ItemList ?? sps?.ItemLists?.FirstOrDefault()?.ItemList ?? "";
                var planningDisplay = (rawPlanning.All(char.IsDigit) && !string.IsNullOrWhiteSpace(spsItem)) ? spsItem : rawPlanning;
                var productInfoDisplay = (rawHose.Equals(rawPlanning, StringComparison.OrdinalIgnoreCase) && sps != null) ? sps.HoseType : rawHose;

                ws.Cells[row, 1].Value = (item.DocumentNumber ?? "").Replace("#", "");
                ws.Cells[row, 2].Value = item.CreatedDate.ToString("dd MMM yyyy HH:mm");
                ws.Cells[row, 3].Value = string.IsNullOrWhiteSpace(planningDisplay) ? "-" : planningDisplay;
                ws.Cells[row, 4].Value = item.CreatedBy ?? "SYSTEM";
                ws.Cells[row, 5].Value = item.MachineName ?? "-";
                ws.Cells[row, 6].Value = string.IsNullOrWhiteSpace(productInfoDisplay) ? "UNDEFINED_PART" : productInfoDisplay;
                ws.Cells[row, 7].Value = item.QtyOk;
                ws.Cells[row, 8].Value = item.NgDimension;
                ws.Cells[row, 9].Value = item.VinCode ?? "-";
                ws.Cells[row, 10].Value = item.Shift ?? "-";
                ws.Cells[row, 11].Value = item.CustomerName ?? "UNKNOWN_CLIENT";

                for (int col = 1; col <= 11; col++)
                {
                    ws.Cells[row, col].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Hair, System.Drawing.Color.FromArgb(148, 163, 184));
                }

                row++;
            }

            ws.Column(1).Width = 24;
            ws.Column(2).Width = 20;
            ws.Column(3).Width = 14;
            ws.Column(4).Width = 16;
            ws.Column(5).Width = 14;
            ws.Column(6).Width = 28;
            ws.Column(7).Width = 12;
            ws.Column(8).Width = 12;
            ws.Column(9).Width = 16;
            ws.Column(10).Width = 10;
            ws.Column(11).Width = 20;

            var fileName = $"DimensionHistory_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            return File(
                package.GetAsByteArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName
            );
        }

        public async Task<IActionResult> Details(int? id)
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("UserName")))
            {
                return RedirectToAction("Login", "Account");
            }
            if (id == null) return NotFound();
            var report = await _context.DimensionReports
                .Include(r => r.Measurements)
                .Include(r => r.Summaries)
                .FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound();

            // Fetch specific SPS record. Match by DocumentNumber (since it refers to the SPS Document)
            ViewBag.Sps = await _context.SpsNoDocs
                .OrderByDescending(s => s.DocumentNumber)
                .FirstOrDefaultAsync(s => s.DocumentNumber == report.DocumentNumber);

            // Fetch Masterlist data as fallback - Match by ItemCode
            ViewBag.Master = await _context.SpsNoDocs
                .OrderByDescending(m => m.DocumentNumber)
                .FirstOrDefaultAsync(m => m.DocumentNumber != null && m.DocumentNumber.Contains(report.ItemCode ?? ""));

            // Fetch SpsItemList shrink data using VinCode / ItemCode
            var vinCode = report.VinCode ?? report.ItemCode ?? "";
            var sanitizedVin = vinCode.Trim().Replace("-", "").Replace(" ", "").ToUpper();
            var itemListEntry = await _context.SpsItemLists
                .FirstOrDefaultAsync(il => il.ItemList != null &&
                    il.ItemList.Replace("-", "").Replace(" ", "").ToUpper() == sanitizedVin);
            // Also check summaries for per-row vin codes if main record doesn't match
            if (itemListEntry == null && report.Summaries != null)
            {
                var firstVin = report.Summaries.FirstOrDefault()?.VinCode;
                if (!string.IsNullOrEmpty(firstVin))
                {
                    var sanVin = firstVin.Trim().Replace("-", "").Replace(" ", "").ToUpper();
                    itemListEntry = await _context.SpsItemLists
                        .FirstOrDefaultAsync(il => il.ItemList != null &&
                            il.ItemList.Replace("-", "").Replace(" ", "").ToUpper() == sanVin);
                }
            }
            ViewBag.ShrinkData = itemListEntry;

            // Build ItemMap for all VINs/Items in report summaries
            var allSpsItems = await _context.SpsItemLists.ToListAsync();
            var itemMap = new Dictionary<string, SpsItemList>(StringComparer.OrdinalIgnoreCase);
            foreach (var il in allSpsItems)
            {
                if (string.IsNullOrEmpty(il.ItemList)) continue;
                var san = il.ItemList.Trim().Replace("-", "").Replace(" ", "").ToUpper();
                if (!itemMap.ContainsKey(san)) itemMap[san] = il;
                if (!itemMap.ContainsKey(il.ItemList)) itemMap[il.ItemList] = il;
            }
            ViewBag.ItemMap = itemMap;

            return View(report);
        }

        public async Task<IActionResult> Edit(int id)
        {
            if (string.IsNullOrEmpty(HttpContext.Session.GetString("UserName")))
            {
                return RedirectToAction("Login", "Account");
            }

            var sessionRole = HttpContext.Session.GetString("UserRole")?.ToUpper() ?? "";
            var isLeaderOrAdmin = sessionRole.Contains("LEADER") || sessionRole.Contains("ADMIN") || sessionRole == "SUPERADMIN";
            if (!isLeaderOrAdmin)
            {
                TempData["ErrorMessage"] = "Akses ditolak: Menu Edit/Revisi data hanya untuk Role LEADER.";
                return RedirectToAction("Details", new { id = id, portal = "dimension" });
            }

            var report = await _context.DimensionReports
                .Include(r => r.Measurements)
                .Include(r => r.Summaries)
                .FirstOrDefaultAsync(r => r.Id == id);
            if (report == null) return NotFound();

            // Fetch specific SPS record. Match by DocumentNumber
            ViewBag.Sps = await _context.SpsNoDocs
                .OrderByDescending(s => s.DocumentNumber)
                .FirstOrDefaultAsync(s => s.DocumentNumber == report.DocumentNumber);

            // Fetch Masterlist data as fallback
            ViewBag.Master = await _context.SpsNoDocs
                .OrderByDescending(m => m.DocumentNumber)
                .FirstOrDefaultAsync(m => (m.DocumentNumber != null && m.DocumentNumber.Contains(report.ItemCode ?? "")) || m.DocumentNumber == report.DocumentNumber);

            return View("Edit", report);
        }

        // AJAX API: Search SPS Items for Autocomplete
        [HttpGet]
        public async Task<IActionResult> SearchSpsItems(string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return Json(new List<string>());
            
            var sanitizedQuery = query.Trim().ToUpper();
            
            // Get unique ItemList values from SpsItemList table
            var spsItems = await _context.SpsItemLists
                .Include(i => i.SpsNoDoc)
                .Where(s => s.ItemList != null && s.ItemList.ToUpper().Contains(sanitizedQuery))
                .Select(s => new { ItemList = s.ItemList, HoseType = s.SpsNoDoc!.HoseType })
                .Distinct()
                .Take(20)
                .ToListAsync();
            
            // Combine and order
            var results = spsItems
                .OrderBy(x => x.ItemList)
                .Take(10)
                .Select(x => new { 
                    itemCode = x.ItemList, 
                    hoseType = x.HoseType ?? "" 
                })
                .ToList();
            
            return Json(results);
        }

        // AJAX API: Get SPS by Item Code (Scanner)
        [HttpGet]
        public async Task<IActionResult> GetSpsByItem(string itemCode)
        {
            if (string.IsNullOrWhiteSpace(itemCode)) return BadRequest();
            
            var sanitizedCode = itemCode.Trim().ToUpper().Replace("-", "");
            
            // Priority 1: Search by SpsItemList.ItemList (kode item dari planning)
            var spsByItemList = await _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .OrderByDescending(s => s.DocumentNumber)
                .FirstOrDefaultAsync(s =>
                    s.ItemLists.Any(il => il.ItemList != null && il.ItemList.Replace("-", "").ToUpper().Contains(sanitizedCode)));

            if (spsByItemList != null) return Json(BuildDimensiSpsResponse(spsByItemList));

            // Priority 2: Search by DocumentNumber
            var spsByDoc = await _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .OrderByDescending(s => s.DocumentNumber)
                .FirstOrDefaultAsync(s =>
                    s.DocumentNumber != null && s.DocumentNumber.Replace("-", "").ToUpper().Contains(sanitizedCode));

            if (spsByDoc != null) return Json(BuildDimensiSpsResponse(spsByDoc));

            // Priority 3: Search by HoseType
            var spsByHose = await _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .OrderByDescending(s => s.DocumentNumber)
                .FirstOrDefaultAsync(s =>
                    s.HoseType != null && s.HoseType.Replace("-", "").ToUpper().Contains(sanitizedCode));

            if (spsByHose != null) return Json(BuildDimensiSpsResponse(spsByHose));

            return Json(new { success = false, message = "Protocol Not Found" });
        }

        private object BuildDimensiSpsResponse(SpsNoDoc s)
        {
            // Helper: build a "Min | Asli | Max" range string from decimal fields
            string BuildRange(decimal? asli, decimal? min, decimal? max, string? fallbackStr = "")
            {
                if (asli.HasValue)
                {
                    if (min.HasValue && max.HasValue)
                        return $"{min:F2} | {asli:F2} | {max:F2}";
                    return asli.Value.ToString("F2");
                }
                return fallbackStr ?? "";
            }

            return new
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
                innerDie = BuildRange(s.Nipple_Asli, s.Nipple_Min, s.Nipple_Max, s.Nipple),
                tubeDie = BuildRange(s.TubeDie_Asli, s.TubeDie_Min, s.TubeDie_Max, s.TubeDie),
                middleDie = BuildRange(s.MiddleDie_Asli, s.MiddleDie_Min, s.MiddleDie_Max, s.MiddleDie),
                coverDie = BuildRange(s.CoverDie_Asli, s.CoverDie_Min, s.CoverDie_Max, s.CoverDie),
                spacerDie = BuildRange(s.SpacerDie_Asli, s.SpacerDie_Min, s.SpacerDie_Max, s.SpacerDie),
                toleranceDie = BuildRange(s.ADistance_Asli, s.ADistance_Min, s.ADistance_Max, s.ADistance),

                // Inner extruder
                headTempInner = BuildRange(s.HeadTemp1_Asli, s.HeadTemp1_Min, s.HeadTemp1_Max, s.HeadTemp1),
                cylinder1TempInner = BuildRange(s.Cylinder1_1_Asli, s.Cylinder1_1_Min, s.Cylinder1_1_Max, s.Cylinder1_1),
                cylinder2TempInner = BuildRange(s.Cylinder2_1_Asli, s.Cylinder2_1_Min, s.Cylinder2_1_Max, s.Cylinder2_1),
                cylinder3TempInner = BuildRange(s.Cylinder3_1_Asli, s.Cylinder3_1_Min, s.Cylinder3_1_Max, s.Cylinder3_1),
                screwTempInner = BuildRange(s.ScrewTemp1_Asli, s.ScrewTemp1_Min, s.ScrewTemp1_Max, s.ScrewTemp1),
                screwSpeedInner = BuildRange(s.ScrewSpeed1_Asli, s.ScrewSpeed1_Min, s.ScrewSpeed1_Max, s.ScrewSpeed1),
                pressureInner = BuildRange(s.Pressure1_Asli, s.Pressure1_Min, s.Pressure1_Max, s.Pressure1),
                feedRollRatioInner = BuildRange(s.FeedRollRatio1_Asli, s.FeedRollRatio1_Min, s.FeedRollRatio1_Max,
                    !string.IsNullOrEmpty(s.FeedRollRatio1) ? s.FeedRollRatio1 : s.Feed1),

                // Outer extruder
                headTempOuter = BuildRange(s.HeadTemp2_Asli, s.HeadTemp2_Min, s.HeadTemp2_Max, s.HeadTemp2),
                cylinder1TempOuter = BuildRange(s.Cylinder1_2_Asli, s.Cylinder1_2_Min, s.Cylinder1_2_Max, s.Cylinder1_2),
                cylinder2TempOuter = BuildRange(s.Cylinder2_2_Asli, s.Cylinder2_2_Min, s.Cylinder2_2_Max, s.Cylinder2_2),
                cylinder3TempOuter = BuildRange(s.Cylinder3_2_Asli, s.Cylinder3_2_Min, s.Cylinder3_2_Max, s.Cylinder3_2),
                screwTempOuter = BuildRange(s.ScrewTemp2_Asli, s.ScrewTemp2_Min, s.ScrewTemp2_Max, s.ScrewTemp2),
                screwSpeedOuter = BuildRange(s.ScrewSpeed2_Asli, s.ScrewSpeed2_Min, s.ScrewSpeed2_Max, s.ScrewSpeed2),
                pressureOuter = BuildRange(s.Pressure2_Asli, s.Pressure2_Min, s.Pressure2_Max, s.Pressure2),
                feedRollRatioOuter = BuildRange(s.FeedRollRatio2_Asli, s.FeedRollRatio2_Min, s.FeedRollRatio2_Max,
                    !string.IsNullOrEmpty(s.FeedRollRatio2) ? s.FeedRollRatio2 : s.Feed2),

                // Process params
                spiralSpeed = BuildRange(s.SpiralSpeed_Asli, s.SpiralSpeed_Min, s.SpiralSpeed_Max, s.SpiralSpeed),
                toleranceSpiralPitch = BuildRange(s.SpiralPitchSetting_Asli, s.SpiralPitchSetting_Min, s.SpiralPitchSetting_Max,
                    !string.IsNullOrEmpty(s.SpiralPitchSetting) ? s.SpiralPitchSetting : s.ToleranceSpiralPitch),
                presetValue = BuildRange(s.PresetValue_Asli, s.PresetValue_Min, s.PresetValue_Max, s.PresetValue),
                controlValue = BuildRange(s.ControlValue_Asli, s.ControlValue_Min, s.ControlValue_Max, s.ControlValue),
                hoseSpeed = BuildRange(s.HoseSpeed_Asli, s.HoseSpeed_Min, s.HoseSpeed_Max, s.HoseSpeed),
                takeupConveyorSpeed = BuildRange(s.TakeUpConveyorSpeed_Asli, s.TakeUpConveyorSpeed_Min, s.TakeUpConveyorSpeed_Max, s.TakeUpConveyorSpeed),
                coolConveyorSpeed = BuildRange(s.CoolConveyorSpeed_Asli, s.CoolConveyorSpeed_Min, s.CoolConveyorSpeed_Max, s.CoolConveyorSpeed),
                conveyorRatio = BuildRange(s.ConveyorRatio_Asli, s.ConveyorRatio_Min, s.ConveyorRatio_Max, s.ConveyorRatio),
                chillerWaterTemp = BuildRange(s.ChillerWaterTemp_Asli, s.ChillerWaterTemp_Min, s.ChillerWaterTemp_Max, s.ChillerWaterTemp),
                caterpillarGap = BuildRange(s.CaterpillarGap_Asli, s.CaterpillarGap_Min, s.CaterpillarGap_Max, s.CaterpillarGap),
                unsmoothSurface = s.UnsmoothSurface ?? "OK",

                // Dimension standards
                toleranceInner = s.ToleranceInner,
                toleranceOuter = s.ToleranceOuter,
                tebalInner = BuildRange(s.TebalInner_Asli, s.TebalInner_Min, s.TebalInner_Max, s.TebalInner),
                tebalTotal = BuildRange(s.TebalTotal_Asli, s.TebalTotal_Min, s.TebalTotal_Max, s.TebalTotal),
                dimensi = s.Dimensi,

                // Dedicated exact targets for Javascript Dimension cards parsing
                // innerTarget = actual inner diameter center value. In DB, ToleranceInner_Asli mistakenly holds the Inner Diameter (e.g. 33.3)
                innerTarget = s.ToleranceInner_Asli?.ToString("F2") ?? (!string.IsNullOrWhiteSpace(s.ThickTarget) ? s.ThickTarget : s.Dimensi),
                innerTol = (s.ToleranceInner_Max.HasValue && s.ToleranceInner_Asli.HasValue) ? (s.ToleranceInner_Max.Value - s.ToleranceInner_Asli.Value).ToString("F2") : s.InnerTol,
                
                // thickTarget = actual inner thickness value. In DB, TebalInner_Asli mistakenly holds the Inner Thickness (e.g. 2.3)
                thickTarget = s.TebalInner_Asli?.ToString("F2") ?? s.InnerTarget,
                thickTol = (s.TebalInner_Max.HasValue && s.TebalInner_Asli.HasValue) ? (s.TebalInner_Max.Value - s.TebalInner_Asli.Value).ToString("F2") : s.ThickTol,
                innerMidTarget = s.TebalInnerMiddle_Asli?.ToString("F2") ?? s.InnerMidTarget,
                innerMidTol = (s.TebalInnerMiddle_Max.HasValue && s.TebalInnerMiddle_Asli.HasValue) ? (s.TebalInnerMiddle_Max.Value - s.TebalInnerMiddle_Asli.Value).ToString("F2") : null,
                tebalInnerMiddle = BuildRange(s.TebalInnerMiddle_Asli, s.TebalInnerMiddle_Min, s.TebalInnerMiddle_Max, null),
                totalTarget = s.TebalTotal_Asli?.ToString("F2") ?? s.TotalTarget,
                totalTol = (s.TebalTotal_Max.HasValue && s.TebalTotal_Asli.HasValue) ? (s.TebalTotal_Max.Value - s.TebalTotal_Asli.Value).ToString("F2") : s.TotalTol,
            };
        }

        [HttpPost]
        public async Task<IActionResult> Delete(int id)
        {
            var report = await _context.DimensionReports
                .Include(r => r.Measurements)
                .FirstOrDefaultAsync(r => r.Id == id);
            if (report != null)
            {
                if (report.Measurements != null) _context.DimensionMeasurements.RemoveRange(report.Measurements);
                _context.DimensionReports.Remove(report);
                await _context.SaveChangesAsync();

                // Notify via SignalR
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");
            }
            return RedirectToAction(nameof(History));
        }

        [HttpPost]
        public async Task<IActionResult> SaveData()
        {
            try {
                using var reader = new StreamReader(Request.Body);
                var body = await reader.ReadToEndAsync();
                var data = JsonSerializer.Deserialize<DimensionReportAppData>(body, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                
                if (data == null) return Json(new { success = false, message = "JSON Error" });
                
                var result = await ProcessSaveInternal(data);
                
                // Notify via SignalR for real-time dashboard updates
                Console.WriteLine($">>> SIGNALR: Broadcasting 'ReceiveUpdate' from Dimensi module at {DateTime.Now}");
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");
                
                return Json(result);
            } catch (Exception ex) {
                return Json(new { success = false, message = "System Error: " + ex.Message });
            }
        }

        private async Task<object> ProcessSaveInternal(DimensionReportAppData data)
        {
            try
            {
                DimensionReport? report = null;
                if (data.ReportId > 0)
                {
                    var sessionRole = HttpContext.Session.GetString("UserRole")?.ToUpper() ?? "";
                    var isLeaderOrAdmin = sessionRole.Contains("LEADER") || sessionRole.Contains("ADMIN") || sessionRole == "SUPERADMIN";
                    if (!isLeaderOrAdmin)
                    {
                        return new { success = false, message = "Akses ditolak: Fitur Edit/Revisi data hanya untuk Role LEADER." };
                    }

                    report = await _context.DimensionReports
                        .Include(r => r.Measurements)
                        .FirstOrDefaultAsync(r => r.Id == data.ReportId);
                }

                if (report == null)
                {
                    report = new DimensionReport
                    {
                        DocumentNumber = !string.IsNullOrEmpty(data.SpsId) ? data.SpsId : ("DIM-" + DateTime.Now.ToString("yyyyMMdd-HHmm")),
                        ProductionDate = DateTime.Now,
                        Status = "ACTIVE",
                        CreatedDate = DateTime.Now,
                        CreatedBy = HttpContext.Session.GetString("UserName") ?? "QC Operator",
                        MachineName = HttpContext.Session.GetString("MachineName"),
                        Shift = ShiftHelper.NormalizeShiftLabel(!string.IsNullOrEmpty(data.Shift) ? data.Shift : GetCurrentShift())
                    };
                    _context.DimensionReports.Add(report);
                    await _context.SaveChangesAsync();
                }

                // Log SpsId before update
                Console.WriteLine($"🔍 DEBUG - Before Update: SpsId = {report.SpsId}, Incoming = {data.SpsId}");
                
                report.HoseType = data.HoseType ?? "";
                report.ItemCode = data.ItemCode ?? "";
                report.SpsId = data.SpsId;
                report.MachineName = data.MachineName ?? report.MachineName;
                report.DimensionDisplay = data.DimensionDisplay ?? "";
                report.VinCode = data.VinCode ?? "";
                report.StandardLength = data.StandardLength ?? "";
                report.StandardShrink = data.StandardShrink ?? "";
                report.ActualLength = data.ActualLength ?? "";
                report.QtyTarget = data.QtyTarget ?? 0;
                report.QtyOk = data.QtyOk ?? 0;
                report.NgDimension = data.NgDimension ?? 0;
                report.NgVisual = data.NgVisual ?? 0;
                report.Remark = data.Remark ?? "";
                report.ByPass = data.ByPass ?? "";
                report.CustomerName = data.CustomerName ?? "-";
                report.Yarn = data.Yarn ?? "";
                report.Shift = ShiftHelper.NormalizeShiftLabel(
                    !string.IsNullOrEmpty(data.Shift) ? data.Shift : (report.Shift ?? GetCurrentShift()));
                
                // Explicitly mark entity as modified to ensure EF tracks changes
                _context.Entry(report).State = Microsoft.EntityFrameworkCore.EntityState.Modified;
                
                // Log SpsId after update
                Console.WriteLine($"💾 DEBUG - After Update: SpsId = {report.SpsId}, State = {_context.Entry(report).State}");

                var existing = _context.DimensionMeasurements.Where(m => m.DimensionReportId == report.Id).ToList();
                if (existing.Any()) _context.DimensionMeasurements.RemoveRange(existing);
                
                if (data.DimensionReadings != null)
                {
                    foreach (var mData in data.DimensionReadings)
                    {
                        if (mData == null) continue;
                        var reading = new DimensionMeasurement
                        {
                            DimensionReportId = report.Id,
                            PointName = mData.PointName ?? "Point",
                            TimeSection = string.IsNullOrWhiteSpace(mData.TimeSection) ? DateTime.Now.ToString("HH:mm") : mData.TimeSection,
                            Frequency = string.IsNullOrWhiteSpace(mData.Frequency) ? "30m Sekali" : mData.Frequency,
                            StandardDimension = mData.StandardDimension ?? "-",
                            Initial = mData.Initial ?? "",
                            R1 = decimal.TryParse(mData.Reading1, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var r1) ? r1 : null,
                            R2 = decimal.TryParse(mData.Reading2, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var r2) ? r2 : null,
                            R3 = decimal.TryParse(mData.Reading3, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var r3) ? r3 : null,
                            R4 = decimal.TryParse(mData.Reading4, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var r4) ? r4 : null,
                            R5 = decimal.TryParse(mData.Reading5, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var r5) ? r5 : null,
                            R6 = decimal.TryParse(mData.Reading6, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var r6) ? r6 : null,
                            Status = string.IsNullOrWhiteSpace(mData.Status) ? "OK" : mData.Status,
                            RecordedTime = DateTime.Now
                        };
                        _context.DimensionMeasurements.Add(reading);
                    }
                }

                // Process Summaries
                var existingSummaries = _context.DimensionSummaries.Where(s => s.DimensionReportId == report.Id).ToList();
                if (existingSummaries.Any()) _context.DimensionSummaries.RemoveRange(existingSummaries);

                if (data.ProductionDataSummaries != null)
                {
                    foreach (var sData in data.ProductionDataSummaries)
                    {
                        if (sData == null) continue;
                        var summary = new DimensionSummary
                        {
                            DimensionReportId = report.Id,
                            PartNumber = sData.PartNumber,
                            VinCode = sData.VinCode,
                            StandardLength = sData.StandardLength,
                            StandardShrink = sData.StandardShrink,
                            ActualLength = sData.ActualLength,
                            QtyTarget = sData.QtyTarget ?? 0,
                            QtyOk = sData.QtyOk ?? 0,
                            NgDimension = sData.NgDimension ?? 0,
                            NgVisual = sData.NgVisual ?? 0
                        };
                        _context.DimensionSummaries.Add(summary);
                    }
                }

                await _context.SaveChangesAsync();
                return new { success = true, reportId = report.Id };
            }
            catch (Exception ex)
            {
                return new { success = false, message = "Internal Error: " + ex.Message };
            }
        }

        [HttpPost]
        public async Task<IActionResult> SubmitReport()
        {
            try {
                using var reader = new StreamReader(Request.Body);
                var body = await reader.ReadToEndAsync();
                var data = JsonSerializer.Deserialize<DimensionReportAppData>(body, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                
                if (data == null) return Json(new { success = false, message = "JSON Error" });

                var saveResult = await ProcessSaveInternal(data);
                if (saveResult is { } sr && (bool)((dynamic)sr).success) 
                {
                    var report = await _context.DimensionReports.FindAsync(data.ReportId);
                    if (report != null)
                    {
                        report.Status = "COMPLETED";
                        await _context.SaveChangesAsync();
                        
                        // Notify via SignalR
                        await _hubContext.Clients.All.SendAsync("ReceiveUpdate");
                    }
                    return Json(new { success = true, redirectUrl = "/Dimensi/Index" });
                }
                return Json(saveResult);
            } catch (Exception ex) {
                return Json(new { success = false, message = "Submit Error: " + ex.Message });
            }
        }
    }
}
