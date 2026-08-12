using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using VelastoProductionSystem.Data;
using VelastoProductionSystem.Helpers;
using VelastoProductionSystem.Models;
using VelastoProductionSystem.Hubs;
using Microsoft.AspNetCore.SignalR;

namespace VelastoProductionSystem.Controllers
{
    public class ProductionReportController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ElwpDbContext _elwpContext;
        private readonly IHubContext<DashboardHub> _hubContext;

        public ProductionReportController(ApplicationDbContext context, ElwpDbContext elwpContext, IHubContext<DashboardHub> hubContext)
        {
            _context = context;
            _elwpContext = elwpContext;
            _hubContext = hubContext;
        }

        // GET: ProductionReport (Monitoring List)
        public async Task<IActionResult> Index()
        {
            
            var reports = await _context.ProductionReports
                .OrderByDescending(p => p.CreatedDate)
                .ToListAsync();
            return View(reports);
        }

        // GET: ProductionReport/ParameterHistory
        public async Task<IActionResult> ParameterHistory(string? startDate = null, string? endDate = null, string? machineFilter = null)
        {
            

            var query = _context.ProductionReports

                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(startDate) && DateTime.TryParse(startDate, out var sDate))
            {
                query = query.Where(p => p.ProductionDate >= sDate.Date);
                ViewBag.StartDate = sDate.ToString("yyyy-MM-dd");
            }
            else
            {
                ViewBag.StartDate = "";
            }

            if (!string.IsNullOrWhiteSpace(endDate) && DateTime.TryParse(endDate, out var eDate))
            {
                query = query.Where(p => p.ProductionDate <= eDate.Date.AddDays(1).AddTicks(-1));
                ViewBag.EndDate = eDate.ToString("yyyy-MM-dd");
            }
            else
            {
                ViewBag.EndDate = "";
            }

            if (!string.IsNullOrWhiteSpace(machineFilter))
            {
                query = query.Where(p => p.MachineName == machineFilter);
                ViewBag.MachineFilter = machineFilter;
            }
            else
            {
                ViewBag.MachineFilter = "";
            }

            // Get Machine List for Dropdown — only valid machine codes (start with EXT)
            ViewBag.MachineList = await _context.ProductionReports
                .Where(p => !string.IsNullOrEmpty(p.MachineName) && p.MachineName.StartsWith("EXT"))
                .Select(p => p.MachineName!)
                .Distinct()
                .OrderBy(m => m)
                .ToListAsync();

            var reports = await query
                .OrderByDescending(p => p.CreatedDate)
                .ToListAsync();
            return View(reports);
        }

        [HttpGet]
        public async Task<IActionResult> ExportParameterHistoryExcel(string? startDate = null, string? endDate = null)
        {
            

            var query = _context.ProductionReports
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(startDate) && DateTime.TryParse(startDate, out var sDate))
                query = query.Where(p => p.ProductionDate >= sDate.Date);

            if (!string.IsNullOrWhiteSpace(endDate) && DateTime.TryParse(endDate, out var eDate))
                query = query.Where(p => p.ProductionDate <= eDate.Date.AddDays(1).AddTicks(-1));

            var reports = await query
                .OrderByDescending(p => p.CreatedDate)
                .ToListAsync();

            using var package = new OfficeOpenXml.ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("Parameter History");

            ws.Cells[1, 1].Value = "PARAMETER HISTORY EXPORT";
            ws.Cells[1, 1, 1, 12].Merge = true;
            ws.Cells[1, 1].Style.Font.Bold = true;
            ws.Cells[1, 1].Style.Font.Size = 14;
            ws.Cells[1, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(15, 23, 42));
            ws.Cells[1, 1].Style.Font.Color.SetColor(System.Drawing.Color.White);

            ws.Cells[2, 1].Value = $"Exported at: {DateTime.Now:dd MMM yyyy HH:mm} | Total rows: {reports.Count}";
            ws.Cells[2, 1, 2, 12].Merge = true;
            ws.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells[2, 1].Style.Font.Color.SetColor(System.Drawing.Color.Gray);

            var headers = new[]
            {
                "Production Date", "Doc Number", "Shift", "Machine", "Item Code", "Hose Type",
                "Dimension", "Yarn", "Inner Material", "Outer Material", "Operator", "Status"
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

            int row = 5;
            foreach (var item in reports)
            {
                var itemCode = item.ItemCode ?? item.HoseType ?? "N/A";

                ws.Cells[row, 1].Value = item.ProductionDate.ToString("dd MMM yyyy");
                ws.Cells[row, 2].Value = item.DocumentNumber;
                ws.Cells[row, 3].Value = item.Shift;
                ws.Cells[row, 4].Value = item.MachineName;
                ws.Cells[row, 5].Value = itemCode;
                ws.Cells[row, 6].Value = item.HoseType;
                ws.Cells[row, 7].Value = item.Dimension;
                ws.Cells[row, 8].Value = item.Yarn;
                ws.Cells[row, 9].Value = item.InnerMaterial;
                ws.Cells[row, 10].Value = item.OuterMaterial;
                ws.Cells[row, 11].Value = item.CreatedBy;
                ws.Cells[row, 12].Value = item.Status;

                for (int col = 1; col <= 12; col++)
                {
                    var cell = ws.Cells[row, col];
                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Hair, System.Drawing.Color.FromArgb(148, 163, 184));
                }

                row++;
            }

            ws.Column(1).Width = 18;
            ws.Column(2).Width = 24;
            ws.Column(3).Width = 10;
            ws.Column(4).Width = 16;
            ws.Column(5).Width = 14;
            ws.Column(6).Width = 28;
            ws.Column(7).Width = 16;
            ws.Column(8).Width = 16;
            ws.Column(9).Width = 20;
            ws.Column(10).Width = 20;
            ws.Column(11).Width = 18;
            ws.Column(12).Width = 14;

            var fileName = $"ParameterHistory_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
            return File(
                package.GetAsByteArray(),
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileName
            );
        }

        // GET: ProductionReport/DimensionHistory
        public async Task<IActionResult> DimensionHistory(string? machineFilter = null)
        {
            

            var query = _context.ProductionReports
                .Where(p => !string.IsNullOrEmpty(p.ActualLength) || !string.IsNullOrEmpty(p.VinCode) || p.QtyOk > 0 || p.NgDimension > 0 || p.NgVisual > 0)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(machineFilter))
            {
                query = query.Where(p => p.MachineName == machineFilter);
                ViewBag.MachineFilter = machineFilter;
            }
            else
            {
                ViewBag.MachineFilter = "";
            }

            // Get Machine List for Dropdown — only valid machine codes (start with EXT)
            ViewBag.MachineList = await _context.ProductionReports
                .Where(p => !string.IsNullOrEmpty(p.MachineName) && p.MachineName.StartsWith("EXT"))
                .Select(p => p.MachineName!)
                .Distinct()
                .OrderBy(m => m)
                .ToListAsync();

            var reports = await query
                .OrderByDescending(p => p.CreatedDate)
                .ToListAsync();
            return View(reports);
        }

        // GET: ProductionReport/App (New Tablet App Interface)
        public async Task<IActionResult> App(int? id)
        {
            

            if (id == null)
            {
                // Empty report for Scanner Mode
                var emptyReport = new ProductionReport 
                { 
                    Id = 0,
                    DocumentNumber = "---",
                    RevisionNumber = 0,
                    ProductionDate = DateTime.Today,
                    CustomerName = "---",
                    HoseType = "---",
                    Dimension = "---",
                    Yarn = "---",
                    Shift = "---",
                    CreatedBy = HttpContext.Session.GetString("UserName") ?? "Operator",
                    MachineName = HttpContext.Session.GetString("MachineName")
                };
                return View(emptyReport);
            }

            var report = await _context.ProductionReports
                .FirstOrDefaultAsync(m => m.Id == id);

            if (report == null) return NotFound();

            // Fetch Masterlist data as fallback standards
            ViewBag.Masterlist = await _context.SpsNoDocs
                .FirstOrDefaultAsync(m => m.HoseType == report.HoseType || m.DocumentNumber == report.HoseType);

            ViewBag.Sps = await _context.SpsNoDocs
                .FirstOrDefaultAsync(s => s.HoseType == report.HoseType || s.DocumentNumber == report.HoseType);

            return View(report);
        }

        // GET: ProductionReport/AppByItem (For Barcode Scanner)
        public async Task<IActionResult> AppByItem(string itemCode)
        {
            var master = await _context.SpsNoDocs
                .FirstOrDefaultAsync(m => m.DocumentNumber == itemCode);
            
            if (master == null) 
            {
                TempData["ErrorMessage"] = $"Item Code '{itemCode}' tidak ditemukan di Database SPS Masterlist!";
                return RedirectToAction(nameof(App));
            }
            
            var sps = await _context.SpsNoDocs
                .FirstOrDefaultAsync(s => s.DocumentNumber == itemCode);

            var report = new ProductionReport {
                Id = 0,
                DocumentNumber = master.DocumentNumber ?? "-",
                RevisionNumber = int.TryParse(master.RevisionNumber, out int rev) ? rev : 0,
                ProductionDate = DateTime.Today,
                CustomerName = master.Customer ?? "-",
                HoseType = master.HoseType ?? "-",
                Dimension = master.Dimensi ?? "-",
                Yarn = master.InnerTube ?? "-",
                Shift = ShiftHelper.GetCurrentShift(await _context.ShiftMasters.ToListAsync()),
                SpsId = null, // legacy id, use DocumentNumber instead
                CreatedBy = HttpContext.Session.GetString("UserName") ?? "Operator"
            };
            
            ViewBag.Masterlist = master;
            TempData["SuccessMessage"] = $"Berhasil memuat spesifikasi untuk '{itemCode}'!";
            
            return View("App", report);
        }

        // AJAX API: Get SPS by ID
        [HttpGet]
        public async Task<IActionResult> GetSpsById(int id)
        {
            var sps = await _context.SpsNoDocs.FindAsync(id);
            if (sps == null) return NotFound();
            return Json(sps);
        }

        // AJAX API: Search SPS Items for Autocomplete (by ItemList OR DocumentNumber)
        [HttpGet]
        public async Task<IActionResult> SearchSpsItems(string query, string? startDate = null, string? endDate = null, string? shift = null, string? hoseType = null, string? pic = null)
        {
            try
            {
                Console.WriteLine($"[SearchSpsItems] === START ===");
                Console.WriteLine($"[SearchSpsItems] Params: query='{query}', startDate='{startDate}', endDate='{endDate}', shift='{shift}', pic='{pic}'");
                
                var sanitizedQuery = string.IsNullOrWhiteSpace(query) ? "" : query.Trim().ToUpper();
                
                // Step 1: Count total DimensionReports in database
                var totalDimensions = await _context.DimensionReports.CountAsync();
                Console.WriteLine($"[SearchSpsItems] Total DimensionReports in DB: {totalDimensions}");
                
                // Step 2: Query DimensionReports with minimal filters
                IQueryable<DimensionReport> dimensionQuery = _context.DimensionReports;

                // Apply date filters only if provided
                if (!string.IsNullOrWhiteSpace(startDate) && DateTime.TryParse(startDate, out DateTime sDate))
                {
                    var beforeFilter = await dimensionQuery.CountAsync();
                    dimensionQuery = dimensionQuery.Where(r => r.CreatedDate >= sDate.Date);
                    var afterFilter = await dimensionQuery.CountAsync();
                    Console.WriteLine($"[SearchSpsItems] Date filter (>= {sDate:yyyy-MM-dd}): {beforeFilter} → {afterFilter}");
                }
                if (!string.IsNullOrWhiteSpace(endDate) && DateTime.TryParse(endDate, out DateTime eDate))
                {
                    var beforeFilter = await dimensionQuery.CountAsync();
                    dimensionQuery = dimensionQuery.Where(r => r.CreatedDate <= eDate.Date.AddDays(1).AddTicks(-1));
                    var afterFilter = await dimensionQuery.CountAsync();
                    Console.WriteLine($"[SearchSpsItems] Date filter (<= {eDate:yyyy-MM-dd}): {beforeFilter} → {afterFilter}");
                }
                if (!string.IsNullOrWhiteSpace(shift) && !shift.Equals("SHIFT ALL", StringComparison.OrdinalIgnoreCase))
                {
                    var shiftNorm = shift.Trim().ToUpper();
                    dimensionQuery = dimensionQuery.Where(r => r.Shift != null && r.Shift.Trim().ToUpper().Contains(shiftNorm));
                }
                if (!string.IsNullOrWhiteSpace(pic))
                {
                    dimensionQuery = dimensionQuery.Where(r => r.CreatedBy != null && r.CreatedBy.ToUpper().Contains(pic.ToUpper()));
                }

                // Load dimension data
                var dimensionData = await dimensionQuery
                    .ToListAsync();

                Console.WriteLine($"[SearchSpsItems] Loaded {dimensionData.Count} dimension reports after filters");
                
                // Check ItemCode distribution
                var withItemCode = dimensionData.Where(d => !string.IsNullOrWhiteSpace(d.ItemCode) || !string.IsNullOrWhiteSpace(d.VinCode)).ToList();
                var withoutItemCode = dimensionData.Count - withItemCode.Count;
                Console.WriteLine($"[SearchSpsItems] ItemCode/VinCode distribution: {withItemCode.Count} have identifiers, {withoutItemCode} are null/empty");
                
                if (withItemCode.Count > 0)
                {
                    Console.WriteLine($"[SearchSpsItems] Sample ItemCodes: {string.Join(", ", withItemCode.Take(10).Select(d => d.ItemCode ?? d.VinCode))}");
                }

                // Map dimension data to search results
                var dimensionItems = withItemCode
                    .Select(d => {
                        var itemCode = (d.ItemCode ?? "").Trim();
                        var vinCode = (d.VinCode ?? "").Trim();
                        var spsItemList = ""; // StandardParameterSetting removed
                        var hoseType = (d.HoseType ?? "").Trim();
                        
                        // Display: prefer SPS ItemList for numeric codes, otherwise use ItemCode as-is (fallback to VinCode if empty)
                        string displayCode;
                        if (!string.IsNullOrWhiteSpace(itemCode) && itemCode.All(char.IsDigit) && !string.IsNullOrWhiteSpace(spsItemList))
                        {
                            displayCode = spsItemList;
                        }
                        else
                        {
                            displayCode = string.IsNullOrWhiteSpace(itemCode) ? vinCode : itemCode;
                        }
                        
                        return new {
                            ItemList = displayCode,
                            OriginalCode = itemCode,
                            VinCode = vinCode,
                            HoseType = hoseType,
                            DocumentNumber = d.DocumentNumber ?? "",
                            CreatedDate = d.CreatedDate
                        };
                    })
                    .Where(x => !string.IsNullOrWhiteSpace(x.DocumentNumber))
                    .ToList();

                Console.WriteLine($"[SearchSpsItems] Mapped to {dimensionItems.Count} items");

                // Apply search filter if provided
                if (!string.IsNullOrEmpty(sanitizedQuery))
                {
                    var beforeSearch = dimensionItems.Count;
                    dimensionItems = dimensionItems
                        .Where(x => 
                            x.DocumentNumber.ToUpper().Contains(sanitizedQuery) ||
                            x.OriginalCode.ToUpper().Contains(sanitizedQuery) ||
                            x.VinCode.ToUpper().Contains(sanitizedQuery) ||
                            x.HoseType.ToUpper().Contains(sanitizedQuery))
                        .ToList();
                    Console.WriteLine($"[SearchSpsItems] After search filter '{sanitizedQuery}': {beforeSearch} → {dimensionItems.Count}");
                }

                // Group by display code and get latest
                var grouped = dimensionItems
                    .GroupBy(x => x.ItemList)
                    .Select(g => g.OrderByDescending(item => item.CreatedDate).First())
                    .ToList();

                Console.WriteLine($"[SearchSpsItems] After grouping: {grouped.Count} unique items");

                // Filter out placeholder values AND numeric codes (incomplete planning)
                var filtered = grouped
                    .Where(x => !IsPlaceholderItemCode(x.ItemList))
                    .ToList();

                Console.WriteLine($"[SearchSpsItems] After placeholder filter: {filtered.Count} items");

                // If still empty, try adding from SPS Master as fallback
                if (filtered.Count == 0)
                {
                    Console.WriteLine($"[SearchSpsItems] No dimension items, trying SPS Master...");
                    var spsFallback2 = await _context.SpsNoDocs
                        .Where(s => s.DocumentNumber != null)
                        .Select(s => new {
                            ItemList = s.DocumentNumber!,
                            OriginalCode = s.DocumentNumber!,
                            VinCode = "",
                            HoseType = s.HoseType ?? "",
                            DocumentNumber = s.DocumentNumber ?? "",
                            CreatedDate = DateTime.MinValue
                        })
                        .ToListAsync();
                    
                    // Filter out placeholder and numeric codes from MasterlistSpsDoubleLayers too
                    spsFallback2 = spsFallback2
                        .Where(x => !IsPlaceholderItemCode(x.DocumentNumber))
                        .Take(50)
                        .ToList();
                    
                    Console.WriteLine($"[SearchSpsItems] MasterlistSpsDoubleLayers fallback (after filter): {spsFallback2.Count} items");
                    filtered = spsFallback2;
                }

                // Final results
                var results = filtered
                    .OrderByDescending(x => x.CreatedDate)
                    .ThenBy(x => x.ItemList)
                    .Take(50)
                    .Select(x => new {
                        itemCode = x.ItemList,
                        hoseType = x.HoseType,
                        documentNumber = x.DocumentNumber
                    })
                    .ToList();

                Console.WriteLine($"[SearchSpsItems] === FINAL RESULTS: {results.Count} items ===");
                if (results.Count > 0)
                {
                    Console.WriteLine($"[SearchSpsItems] Sample results: {string.Join(", ", results.Take(5).Select(r => r.itemCode))}");
                }
                
                return Json(results);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SearchSpsItems ERROR] {ex.Message}");
                Console.WriteLine($"[SearchSpsItems STACK] {ex.StackTrace}");
                return Json(new List<object>());
            }
        }

        // AJAX API: Get SPS by Item Code (Scanner)
        [HttpGet]
        public async Task<IActionResult> GetSpsByItem(string itemCode, string? machineCode = null)
        {
            if (string.IsNullOrWhiteSpace(itemCode)) return BadRequest();
            
            var sanitizedCode = itemCode.Trim().ToUpper().Replace("-", "");
            var sanitizedMachine = machineCode?.Trim().ToUpper();
            
            // Use SpsNoDocs for SPS lookup
            var spsQuery = _context.SpsNoDocs.AsQueryable();
            
            // Filter by ItemCode (Normalized)
            spsQuery = spsQuery.Where(s =>
                s.ItemLists.Any(il => il.ItemList != null && il.ItemList.Replace("-", "").ToUpper().Contains(sanitizedCode)) ||
                (s.DocumentNumber != null && s.DocumentNumber.Replace("-", "").ToUpper().Contains(sanitizedCode)));

            if (!string.IsNullOrEmpty(sanitizedMachine))
            {
                // Detect Machine Category (DL vs CHS)
                bool isDL = sanitizedMachine.Contains("DL") || sanitizedMachine.Contains("DOUBLE") || sanitizedMachine.Contains("NON-CHS");
                bool isCHS = sanitizedMachine.Contains("CHS") && !isDL;

                // 1. Try EXACT machine match
                var spsExact = await spsQuery.Where(s =>
                    (s.MachineCode != null && s.MachineCode.ToUpper() == sanitizedMachine) ||
                    (s.Machine != null && s.Machine.ToUpper() == sanitizedMachine))
                    .OrderByDescending(s => s.DocumentNumber).FirstOrDefaultAsync();
                if (spsExact != null) return Json(MapMasterToSps(spsExact, itemCode.Trim()));

                // 2. Try CATEGORY match (DL to DL, CHS to CHS)
                var spsCategory = await spsQuery.Where(s =>
                    (isDL && s.Machine != null && (s.Machine.ToUpper().Contains("DL") || s.Machine.ToUpper().Contains("NON-CHS") || s.Machine.ToUpper().Contains("DOUBLE"))) ||
                    (isCHS && s.Machine != null && s.Machine.ToUpper().Contains("CHS") && !s.Machine.ToUpper().Contains("DL")))
                    .OrderByDescending(s => s.DocumentNumber).FirstOrDefaultAsync();
                if (spsCategory != null) return Json(MapMasterToSps(spsCategory, itemCode.Trim()));
            }

            // ULTIMATE FALLBACK
            var ultimateFallback = await spsQuery.OrderByDescending(s => s.DocumentNumber).FirstOrDefaultAsync();
            if (ultimateFallback != null) return Json(MapMasterToSps(ultimateFallback, itemCode.Trim()));

            return Json(new { success = false, message = "Protocol Not Found" });
        }

        private object MapMasterToSps(SpsNoDoc master, string? requestedItemCode = null)
        {
            // Helper: build "Min | Asli | Max" range from decimal fields, fall back to string
            static string R(decimal? asli, decimal? min, decimal? max, string? fallback = "")
            {
                if (asli.HasValue)
                {
                    string sMin = min.HasValue ? min.Value.ToString("F2") : "--";
                    string sMax = max.HasValue ? max.Value.ToString("F2") : "--";
                    return $"{sMin} | {asli.Value:F2} | {sMax}";
                }
                
                string fb = fallback?.Trim() ?? "";
                if (string.IsNullOrEmpty(fb) || fb == "-" || fb == "---") return "";
                if (fb.Contains("|")) return fb;
                
                return $"-- | {fb} | --";
            }

            // Resolve item code: prefer the explicitly requested code (from planning/search),
            // then fall back to the first ItemList linked to this SPS document.
            var resolvedItemCode = !string.IsNullOrWhiteSpace(requestedItemCode)
                ? requestedItemCode
                : master.ItemLists?.FirstOrDefault(il => !string.IsNullOrWhiteSpace(il.ItemList))?.ItemList;

            return new {
                // Masterlist SPS rows do not have a numeric StandardParameterSettings FK.
                // Keep the document number separately and leave Id null so clients do not post it into SpsId.
                Id = (string?)null,
                Source = "Masterlist",
                MasterlistId = master.DocumentNumber,
                ItemList = resolvedItemCode,
                ItemCode = resolvedItemCode,
                HoseType = master.HoseType,
                CustomerName = master.Customer,
                DocumentNumber = master.DocumentNumber,
                RevisionNumber = master.RevisionNumber,
                InnerMaterial = master.InnerTube,
                MiddleMaterial = master.MiddleTube,
                OuterMaterial = master.OuterCover,
                YarnType = master.Yarn,
                LayerType = !string.IsNullOrWhiteSpace(master.MiddleTube) ? "CHS 3 Layer" : "CHS 2 Layer",
                
                HeadTempInner      = R(master.HeadTemp1_Asli, master.HeadTemp1_Min, master.HeadTemp1_Max, master.HeadTemp1),
                Cylinder1TempInner = R(master.Cylinder1_1_Asli, master.Cylinder1_1_Min, master.Cylinder1_1_Max, master.Cylinder1_1),
                Cylinder2TempInner = R(master.Cylinder2_1_Asli, master.Cylinder2_1_Min, master.Cylinder2_1_Max, master.Cylinder2_1),
                Cylinder3TempInner = R(master.Cylinder3_1_Asli, master.Cylinder3_1_Min, master.Cylinder3_1_Max, master.Cylinder3_1),
                ScrewTempInner     = R(master.ScrewTemp1_Asli, master.ScrewTemp1_Min, master.ScrewTemp1_Max, master.ScrewTemp1),
                ScrewSpeedInner    = R(master.ScrewSpeed1_Asli, master.ScrewSpeed1_Min, master.ScrewSpeed1_Max, master.ScrewSpeed1),
                PressureInner      = R(master.Pressure1_Asli, master.Pressure1_Min, master.Pressure1_Max, master.Pressure1),
                FeedRollRatioInner = R(master.FeedRollRatio1_Asli, master.FeedRollRatio1_Min, master.FeedRollRatio1_Max,
                    !string.IsNullOrEmpty(master.FeedRollRatio1) ? master.FeedRollRatio1 : master.Feed1),
                
                HeadTempOuter      = R(master.HeadTemp2_Asli, master.HeadTemp2_Min, master.HeadTemp2_Max, master.HeadTemp2),
                Cylinder1TempOuter = R(master.Cylinder1_2_Asli, master.Cylinder1_2_Min, master.Cylinder1_2_Max, master.Cylinder1_2),
                Cylinder2TempOuter = R(master.Cylinder2_2_Asli, master.Cylinder2_2_Min, master.Cylinder2_2_Max, master.Cylinder2_2),
                Cylinder3TempOuter = R(master.Cylinder3_2_Asli, master.Cylinder3_2_Min, master.Cylinder3_2_Max, master.Cylinder3_2),
                ScrewTempOuter     = R(master.ScrewTemp2_Asli, master.ScrewTemp2_Min, master.ScrewTemp2_Max, master.ScrewTemp2),
                ScrewSpeedOuter    = R(master.ScrewSpeed2_Asli, master.ScrewSpeed2_Min, master.ScrewSpeed2_Max, master.ScrewSpeed2),
                PressureOuter      = R(master.Pressure2_Asli, master.Pressure2_Min, master.Pressure2_Max, master.Pressure2),
                FeedRollRatioOuter = R(master.FeedRollRatio2_Asli, master.FeedRollRatio2_Min, master.FeedRollRatio2_Max,
                    !string.IsNullOrEmpty(master.FeedRollRatio2) ? master.FeedRollRatio2 : master.Feed2),
                
                NippleDie    = R(master.Nipple_Asli, master.Nipple_Min, master.Nipple_Max, master.Nipple),
                TubeDie      = R(master.TubeDie_Asli, master.TubeDie_Min, master.TubeDie_Max, master.TubeDie),
                MiddleDie    = R(master.MiddleDie_Asli, master.MiddleDie_Min, master.MiddleDie_Max, master.MiddleDie),
                CoverDie     = R(master.CoverDie_Asli, master.CoverDie_Min, master.CoverDie_Max, master.CoverDie),
                SpacerDie    = R(master.SpacerDie_Asli, master.SpacerDie_Min, master.SpacerDie_Max, master.SpacerDie),
                ToleranceDie = R(master.ADistance_Asli, master.ADistance_Min, master.ADistance_Max, master.ADistance),
                
                ChillerWaterTemp    = R(master.ChillerWaterTemp_Asli, master.ChillerWaterTemp_Min, master.ChillerWaterTemp_Max, master.ChillerWaterTemp),
                TakeupConveyorSpeed = R(master.TakeUpConveyorSpeed_Asli, master.TakeUpConveyorSpeed_Min, master.TakeUpConveyorSpeed_Max, master.TakeUpConveyorSpeed),
                SpiralSpeed         = R(master.SpiralSpeed_Asli, master.SpiralSpeed_Min, master.SpiralSpeed_Max, master.SpiralSpeed),
                ToleranceSpiralPitch = R(master.SpiralPitchSetting_Asli, master.SpiralPitchSetting_Min, master.SpiralPitchSetting_Max,
                    !string.IsNullOrEmpty(master.SpiralPitchSetting) ? master.SpiralPitchSetting : master.ToleranceSpiralPitch),
                SpiralPitchDisplay  = R(master.SpiralPitchDisplay_Asli, master.SpiralPitchDisplay_Min, master.SpiralPitchDisplay_Max, master.SpiralPitchDisplay),
                HoseSpeed           = R(master.HoseSpeed_Asli, master.HoseSpeed_Min, master.HoseSpeed_Max, master.HoseSpeed),
                CoolConveyorSpeed   = R(master.CoolConveyorSpeed_Asli, master.CoolConveyorSpeed_Min, master.CoolConveyorSpeed_Max, master.CoolConveyorSpeed),
                ConveyorRatio       = R(master.ConveyorRatio_Asli, master.ConveyorRatio_Min, master.ConveyorRatio_Max, master.ConveyorRatio),
                PresetValue         = R(master.PresetValue_Asli, master.PresetValue_Min, master.PresetValue_Max, master.PresetValue),
                ControlValue        = R(master.ControlValue_Asli, master.ControlValue_Min, master.ControlValue_Max, master.ControlValue),
                UnsmoothSurface     = master.UnsmoothSurface,
                CaterpillarGap      = R(master.CaterpillarGap_Asli, master.CaterpillarGap_Min, master.CaterpillarGap_Max, master.CaterpillarGap),

                MeshScreen1 = master.MeshScreen1,
                MeshScreen2 = master.MeshScreen2,
                MeshScreen3 = master.MeshScreen3,
                MarkingStandard = master.TextMarkingMaterial
            };
        }

        private static string? SanitizeLegacySpsId(string? spsId)
        {
            if (string.IsNullOrWhiteSpace(spsId)) return null;
            var trimmed = spsId.Trim();
            return int.TryParse(trimmed, out _) ? trimmed : null;
        }

        private static string NormalizeLookupCode(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            return System.Text.RegularExpressions.Regex.Replace(value.Trim(), @"[^A-Za-z0-9]", "").ToUpperInvariant();
        }

        private async Task<SpsNoDoc?> FindSpsNoDocForReportAsync(ProductionReport report)
        {
            var itemNorm = NormalizeLookupCode(report.ItemCode);
            var docNumber = report.DocumentNumber?.Trim();
            var machineUpper = report.MachineName?.Trim().ToUpperInvariant();

            async Task<SpsNoDoc?> PickBestMatch(IQueryable<SpsNoDoc> query)
            {
                if (!string.IsNullOrEmpty(machineUpper))
                {
                    var isDL = machineUpper.Contains("DL") || machineUpper.Contains("DOUBLE") || machineUpper.Contains("NON-CHS");
                    var isCHS = machineUpper.Contains("CHS") && !isDL;

                    var exactMachine = await query
                        .Where(s =>
                            (s.MachineCode != null && s.MachineCode.ToUpper() == machineUpper) ||
                            (s.Machine != null && s.Machine.ToUpper() == machineUpper))
                        .OrderByDescending(s => s.DocumentNumber)
                        .FirstOrDefaultAsync();
                    if (exactMachine != null) return exactMachine;

                    var category = await query
                        .Where(s =>
                            (isDL && s.Machine != null && (s.Machine.ToUpper().Contains("DL") || s.Machine.ToUpper().Contains("NON-CHS") || s.Machine.ToUpper().Contains("DOUBLE"))) ||
                            (isCHS && s.Machine != null && s.Machine.ToUpper().Contains("CHS") && !s.Machine.ToUpper().Contains("DL")))
                        .OrderByDescending(s => s.DocumentNumber)
                        .FirstOrDefaultAsync();
                    if (category != null) return category;
                }

                return await query.OrderByDescending(s => s.DocumentNumber).FirstOrDefaultAsync();
            }

            IQueryable<SpsNoDoc> ByNormalizedCode(string normalized) =>
                _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .Where(s =>
                        s.ItemLists.Any(il => il.ItemList != null && il.ItemList.Replace("-", "").Replace(" ", "").ToUpper().Contains(normalized)) ||
                        (s.DocumentNumber != null && s.DocumentNumber.Replace("-", "").Replace(" ", "").ToUpper().Contains(normalized)));

            // 1) Match by ItemCode through SpsItemLists (same strategy as GetSpsByItem)
            if (!string.IsNullOrEmpty(itemNorm))
            {
                var byItem = await PickBestMatch(ByNormalizedCode(itemNorm));
                if (byItem != null) return byItem;
            }

            // 2) Exact document number from report header
            if (!string.IsNullOrWhiteSpace(docNumber))
            {
                var exactDoc = await _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .FirstOrDefaultAsync(s => s.DocumentNumber == docNumber);
                if (exactDoc != null) return exactDoc;

                var docNorm = NormalizeLookupCode(docNumber);
                if (!string.IsNullOrEmpty(docNorm) && docNorm != itemNorm)
                {
                    var byDoc = await PickBestMatch(ByNormalizedCode(docNorm));
                    if (byDoc != null) return byDoc;
                }
            }

            // 3) Legacy fallback (previous behavior)
            if (!string.IsNullOrWhiteSpace(report.ItemCode))
            {
                return await _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .Where(m => m.DocumentNumber != null && m.DocumentNumber.Contains(report.ItemCode))
                    .OrderByDescending(m => m.DocumentNumber)
                    .FirstOrDefaultAsync();
            }

            return null;
        }

        // GET: ProductionReport/GetTodaysSummary
        // Returns today's submitted data for the operator's machine (used by FAB popup, read-only)
        [HttpGet]
        public async Task<IActionResult> GetTodaysSummary(DateTime? targetDate = null)
        {
            var machineName = HttpContext.Session.GetString("MachineName");
            var isAdmin = HttpContext.Session.GetString("IsAdmin") == "true";
            
            var startDate = (targetDate ?? DateTime.Today).Date;
            var endDate = startDate.AddDays(1);

            // Query ProductionReports for target date
            var paramQuery = _context.ProductionReports
                .Where(r => r.CreatedDate >= startDate && r.CreatedDate < endDate);

            // Query DimensionReports for target date
            var dimQuery = _context.DimensionReports
                .Where(r => r.CreatedDate >= startDate && r.CreatedDate < endDate);

            // Filter by machine for non-admin
            if (!isAdmin && !string.IsNullOrEmpty(machineName))
            {
                paramQuery = paramQuery.Where(r => r.MachineName != null && r.MachineName.Contains(machineName));
                dimQuery   = dimQuery.Where(r => r.MachineName != null && r.MachineName.Contains(machineName));
            }

            var paramReports = await paramQuery
                .OrderByDescending(r => r.CreatedDate)
                .Select(r => new {
                    r.Id, r.HoseType, r.VinCode, r.Shift, r.CreatedBy, r.Status, r.CreatedDate,
                    r.DocumentNumber, r.RevisionNumber,
                    r.InnerMaterialActual, r.InnerMaterialLotNo,
                    r.OuterMaterialActual, r.OuterMaterialLotNo,
                    r.YarnActual, r.YarnLotNo,
                    r.InitHoseSpeed, r.EmbossMarkContent,
                    r.NippleDieOK, r.TubeDieOK, r.MiddleDieOK, r.CoverDieOK, r.SpacerDieOK, r.ToleranceDieOK,
                    r.StandardLength, r.ActualLength,
                    r.QtyTarget, r.QtyOk, r.NgDimension, r.NgVisual, r.Remark,
                    r.MachineName
                }).ToListAsync();

            var dimReports = await dimQuery
                .OrderByDescending(r => r.CreatedDate)
                .Select(r => new {
                    r.Id, r.HoseType, r.VinCode, r.Shift, r.CreatedBy, r.Status,
                    r.DocumentNumber, r.DimensionDisplay,
                    r.StandardLength, r.ActualLength,
                    r.QtyTarget, r.QtyOk, r.NgDimension, r.NgVisual, r.Remark,
                    r.MachineName
                }).ToListAsync();

            return Json(new {
                machine = machineName ?? (isAdmin ? "ALL" : "—"),
                parameterReports = paramReports,
                dimensionReports = dimReports
            });
        }

        [HttpGet]
        public async Task<IActionResult> GetPlanningItems(DateTime date, string shift, string? machine = null)

        {
            var cultureID = new System.Globalization.CultureInfo("id-ID");
            var cultureEN = new System.Globalization.CultureInfo("en-US");
            
            // Ambil nama mesin dari session jika tidak di-pass via parameter (non-admin)
            var sessionMachine = HttpContext.Session.GetString("MachineName");
            var isAdmin = HttpContext.Session.GetString("IsAdmin") == "true";
            var machineFilter = machine ?? sessionMachine;
            
            // 1. PRIORITAS: Ambil Live dari ELWP_PRD (Source of Truth Utama)
            var start = date.Date;
            var end = start.AddDays(1);

            // SELALU filter hanya mesin Extrude (AreaId == 1) - berlaku untuk Admin maupun Operator
            var extrudeMachineIds = await _elwpContext.ElwpMachines
                .Where(m => m.IsActive && m.AreaId == 1)
                .Select(m => m.Id)
                .ToListAsync();

            var elwpQuery = from p in _elwpContext.ElwpPlannings
                             join m in _elwpContext.ElwpMachines on p.MesinId equals m.Id into machineJoin
                             from m in machineJoin.DefaultIfEmpty()
                             where p.TanggalPlanning >= start && p.TanggalPlanning < end
                                   && extrudeMachineIds.Contains(p.MesinId ?? 0)
                             select new { p, MachineName = m != null ? m.NamaMesin : "UNKNOWN", MachineCode = m != null ? m.KodeMesin : "" };

            var elwpRows = await elwpQuery.ToListAsync();

            if (elwpRows.Any())
            {
                // Re-enable shift filtering based on user request
                if (!string.IsNullOrEmpty(shift))
                {
                    elwpRows = elwpRows.Where(x => x.p.Shift != null && x.p.Shift.Equals(shift, StringComparison.OrdinalIgnoreCase)).ToList();
                }

                // Filter berdasarkan mesin spesifik - hanya untuk non-admin
                // jika operator memiliki session mesin, hanya tampilkan planning untuk mesin tersebut
                if (!isAdmin && !string.IsNullOrEmpty(machineFilter))
                {
                    var normalizedFilter = new string(machineFilter.Where(char.IsLetterOrDigit).ToArray()).ToUpper();
                    elwpRows = elwpRows.Where(x =>
                        (x.MachineName != null && new string(x.MachineName.Where(char.IsLetterOrDigit).ToArray()).ToUpper().Contains(normalizedFilter)) ||
                        (x.MachineCode != null && new string(x.MachineCode.Where(char.IsLetterOrDigit).ToArray()).ToUpper().Contains(normalizedFilter))
                    ).ToList();
                }

                // Jangan tampilkan DANDORI di Planning Master form parameter
                // (data tetap ada di database, hanya tidak ditampilkan di dropdown)
                elwpRows = elwpRows
                    .Where(x => !string.Equals(x.p.KodeItem?.Trim(), "DANDORI", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                // Batch-check SPS availability for all item codes
                var elwpItemCodes = elwpRows
                    .Where(x => !string.IsNullOrEmpty(x.p.KodeItem))
                    .Select(x => x.p.KodeItem!.Trim().ToUpper())
                    .Distinct().ToList();

                var spsItemCodes = await _context.SpsItemLists
                    .Where(s => s.ItemList != null)
                    .Select(s => s.ItemList!.ToUpper())
                    .ToListAsync();

                var masterItemCodes = await _context.SpsNoDocs
                    .Where(m => m.DocumentNumber != null)
                    .Select(m => m.DocumentNumber!.ToUpper())
                    .ToListAsync();

                bool HasSps(string? code) {
                    if (string.IsNullOrEmpty(code)) return false;
                    var u = code.Trim().ToUpper();
                    return spsItemCodes.Any(s => s.Contains(u)) || masterItemCodes.Any(s => s.Contains(u));
                }

                var items = elwpRows.Select(x => new {
                    itemCode = x.p.KodeItem,
                    itemName = (x.p.PartName ?? "#N/A") + (string.IsNullOrEmpty(x.p.PnSap) ? "" : " | " + x.p.PnSap),
                    machineName = x.MachineName,
                    machineCode = x.MachineCode,
                    dateShift = $"{(x.p.TanggalPlanning?.ToString("dddd, d MMMM yyyy", cultureID) ?? "").ToUpper()} SHIFT {x.p.Shift}",
                    date = x.p.TanggalPlanning?.ToString("yyyy-MM-dd"),
                    shift = x.p.Shift,
                    isFiltered = !isAdmin && !string.IsNullOrEmpty(machineFilter),
                    hasSps = HasSps(x.p.KodeItem)
                }).OrderBy(p => p.itemName).ToList();

                return Json(items);
            }

            // 2. FALLBACK: Jika di ELWP kosong (data manual/lama), ambil dari Cache Lokal
            var dayStr1 = date.Day.ToString();
            var dayStr2 = date.Day.ToString("00");
            var months = new List<string> {
                date.ToString("MMMM", cultureID).ToUpper(), date.ToString("MMM", cultureID).ToUpper(),
                date.ToString("MMMM", cultureEN).ToUpper(), date.ToString("MMM", cultureEN).ToUpper()
            }.Distinct().ToList();

            var yearFull = date.Year.ToString();
            var yearShort = date.ToString("yy");

            var itemRows = await _context.PlanningMasters
                .Where(p => p.DateShiftString != null)
                .Select(p => new { p.Id, p.PartName1, p.PartName2, p.Kode, p.DateShiftString })
                .ToListAsync();

            var filtered = itemRows.Where(p => {
                var ds = (p.DateShiftString ?? "").ToUpper();
                bool hasDayMatch = ds.Contains(dayStr1) || ds.Contains(dayStr2);
                bool hasMonthMatch = months.Any(m => ds.Contains(m));
                bool hasYearMatch = ds.Contains(yearFull) || ds.Contains(yearShort);
                bool hasShiftMatch = string.IsNullOrEmpty(shift) || ds.Contains(shift.ToUpper());
                return hasDayMatch && hasMonthMatch && hasYearMatch && hasShiftMatch;
            }).OrderByDescending(x => x.Id).ToList();

            if (!filtered.Any())
            {
                filtered = itemRows.OrderByDescending(x => x.Id).Take(50).ToList();
            }

            // Batch-check SPS for fallback items (must be outside Select)
            var fallbackCodes = filtered
                .Select(p => !string.IsNullOrWhiteSpace(p.Kode) ? p.Kode : p.PartName1)
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c!.Trim().ToUpper())
                .Distinct().ToList();

            var fbSpsCodes = await _context.SpsItemLists
                .Where(s => s.ItemList != null)
                .Select(s => s.ItemList!.ToUpper())
                .ToListAsync();

            var fbMasterCodes = await _context.SpsNoDocs
                .Where(m => m.DocumentNumber != null)
                .Select(m => m.DocumentNumber!.ToUpper())
                .ToListAsync();

            bool FbHasSps(string? code) {
                if (string.IsNullOrEmpty(code)) return false;
                var u = code.Trim().ToUpper();
                return fbSpsCodes.Any(s => s.Contains(u)) || fbMasterCodes.Any(s => s.Contains(u));
            }

            var result = filtered.Select(p => {
                var parsed = ParseDateShiftString(p.DateShiftString);
                // BUGFIX: PartName1 adalah KodeItem asli (TA1670, NA1860), Kode adalah PnSap
                // Prioritaskan PartName1, fallback ke Kode hanya jika PartName1 kosong
                var fallbackItemCode = !string.IsNullOrWhiteSpace(p.PartName1) ? p.PartName1 : p.Kode;
                return new {
                    itemCode = fallbackItemCode,
                    itemName = (p.PartName2 ?? "") + (string.IsNullOrEmpty(p.Kode) ? "" : " | " + p.Kode),
                    dateShift = p.DateShiftString,
                    date = parsed.date,
                    shift = parsed.shift,
                    hasSps = FbHasSps(fallbackItemCode)
                };
            }).Distinct().OrderBy(p => p.itemName).ToList();

            return Json(result);

        }

        private static (string date, string shift) ParseDateShiftString(string? dateShiftString)
        {
            if (string.IsNullOrWhiteSpace(dateShiftString)) return (string.Empty, string.Empty);

            var normalized = dateShiftString.ToUpper().Trim();
            var parts = normalized.Split(new[] { " SHIFT " }, StringSplitOptions.None);
            if (parts.Length != 2) return (string.Empty, string.Empty);

            var dayString = parts[0];
            var shiftString = parts[1].Trim();

            // Hapus hari nama di depan, contoh: "RABU, 8 APRIL 2026" -> "8 APRIL 2026"
            var commaIndex = dayString.IndexOf(',');
            if (commaIndex >= 0 && commaIndex + 1 < dayString.Length)
            {
                dayString = dayString.Substring(commaIndex + 1).Trim();
            }

            var dateParts = dayString.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (dateParts.Length < 3) return (string.Empty, string.Empty);

            var day = dateParts[0].PadLeft(2, '0');
            var month = dateParts[1];
            var year = dateParts[2];

            var monthMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "JANUARI", "01" }, { "FEBRUARI", "02" }, { "MARET", "03" }, { "APRIL", "04" },
                { "MEI", "05" }, { "JUNI", "06" }, { "JULI", "07" }, { "AGUSTUS", "08" },
                { "SEPTEMBER", "09" }, { "OKTOBER", "10" }, { "NOVEMBER", "11" }, { "DESEMBER", "12" }
            };

            if (!monthMap.TryGetValue(month, out var monthNumber))
                monthNumber = "01";

            var formattedDate = $"{year}-{monthNumber}-{day}";

            string normalizedShift = ShiftHelper.NormalizeShiftLabel(shiftString);

            return (formattedDate, normalizedShift);
        }

        // GET: ProductionReport/Details/5 (The Monitoring Dashboard - Gambar 2 & 5)
        [HttpPost]
        public async Task<IActionResult> SaveWithReadings([FromBody] ProductionReportSaveDto dto)
        {
            if (dto == null) return Json(new { success = false, message = "Invalid data payload" });

            try
            {
                var machineName = HttpContext.Session.GetString("MachineName");
                var (productionStart, productionEnd) = await GetProductionWindowAsync(machineName);

                var report = new ProductionReport
                {
                    DocumentNumber = dto.DocumentNumber ?? "AUTO",
                    RevisionNumber = dto.RevisionNumber,
                    ProductionDate = DateTime.TryParse(dto.ProductionDate, out var dt) ? dt : DateTime.Now,
                    Shift = ShiftHelper.NormalizeShiftLabel(dto.Shift),
                    CustomerName = dto.CustomerName,
                    HoseType = dto.HoseType,
                    MachineName = machineName,
                    InnerMaterial = dto.InnerMaterial,
                    InnerMaterialActual = dto.InnerMaterialActual,
                    OuterMaterial = dto.OuterMaterial,
                    OuterMaterialActual = dto.OuterMaterialActual,
                    MiddleMaterial = dto.MiddleMaterial,
                    MiddleMaterialActual = dto.MiddleMaterialActual,
                    Yarn = dto.Yarn,
                    YarnActual = dto.YarnActual,

                    // AUTO-MAPPING: Ambil Lot & SG pertama dari list untuk ditampilkan di Identity Card
                    InnerMaterialLotNo = dto.MaterialLots?.FirstOrDefault(x => x.LayerType == "INNER")?.LotNumber,
                    InnerMaterialSG = dto.MaterialLots?.FirstOrDefault(x => x.LayerType == "INNER")?.SGValue,
                    OuterMaterialLotNo = dto.MaterialLots?.FirstOrDefault(x => x.LayerType == "OUTER")?.LotNumber,
                    OuterMaterialSG = dto.MaterialLots?.FirstOrDefault(x => x.LayerType == "OUTER")?.SGValue,
                    MiddleMaterialLotNo = dto.MaterialLots?.FirstOrDefault(x => x.LayerType == "MIDDLE")?.LotNumber,
                    MiddleMaterialSG = dto.MaterialLots?.FirstOrDefault(x => x.LayerType == "MIDDLE")?.SGValue,
                    YarnLotNo = dto.MaterialLots?.FirstOrDefault(x => x.LayerType == "YARN")?.LotNumber,

                    DAI_Awal = dto.DAI_Awal,
                    DAI_Akhir = dto.DAI_Akhir,
                    DAC_Awal = dto.DAC_Awal,
                    DAC_Akhir = dto.DAC_Akhir,
                    DRAI_Awal = dto.DRAI_Awal,
                    DRAI_Akhir = dto.DRAI_Akhir,
                    DRAC_Awal = dto.DRAC_Awal,
                    DRAC_Akhir = dto.DRAC_Akhir,
                    CreatedBy = dto.CreatedBy ?? "Operator",
                    CreatedDate = DateTime.Now,
                    ProductionStartTime = productionStart,
                    ProductionEndTime = productionEnd,
                    Status = "COMPLETED",
                    SpsId = SanitizeLegacySpsId(dto.SpsId),
                    ItemCode = dto.ItemCode,
                    NippleDieOK = dto.NippleDieOK, NippleDieInitial = dto.NippleDieInitial, NippleDieFinal = dto.NippleDieFinal,
                    TubeDieOK = dto.TubeDieOK, TubeDieInitial = dto.TubeDieInitial, TubeDieFinal = dto.TubeDieFinal,
                    MiddleDieOK = dto.MiddleDieOK, MiddleDieInitial = dto.MiddleDieInitial, MiddleDieFinal = dto.MiddleDieFinal,
                    CoverDieOK = dto.CoverDieOK, CoverDieInitial = dto.CoverDieInitial, CoverDieFinal = dto.CoverDieFinal,
                    SpacerDieOK = dto.SpacerDieOK, SpacerDieInitial = dto.SpacerDieInitial, SpacerDieFinal = dto.SpacerDieFinal,
                    ToleranceDieOK = dto.ToleranceDieOK, ToleranceInitial = dto.ToleranceInitial, ToleranceFinal = dto.ToleranceFinal,
                    MeshInner10Before = dto.MeshInner10Before, MeshInner40Before = dto.MeshInner40Before,
                    MeshOuter10Before = dto.MeshOuter10Before, MeshOuter40Before = dto.MeshOuter40Before,
                    MeshInnerCheck = dto.MeshInnerCheck, MeshOuterCheck = dto.MeshOuterCheck,
                    EmbossMarkContent = dto.EmbossMarkContent,
                    EmbossMarkDate = DateTime.TryParse(dto.EmbossMarkDate, out var ed) ? ed : null,
                    QcCond = dto.QcCond, QcSurf = dto.QcSurf, QcRes = dto.QcRes
                };

                // Sync Initial instrument parameters from the first reading if they are provided
                if (dto.Readings != null && dto.Readings.Count > 0)
                {
                    var r0 = dto.Readings[0];
                    report.InitHeadTempInner = ParseDecimal(r0.HeadTempInner);
                    report.InitCylinder1TempInner = ParseDecimal(r0.Cylinder1TempInner);
                    report.InitCylinder2TempInner = ParseDecimal(r0.Cylinder2TempInner);
                    report.InitCylinder3TempInner = ParseDecimal(r0.Cylinder3TempInner);
                    report.InitScrewTempInner = ParseDecimal(r0.ScrewTempInner);
                    report.InitScrewSpeedInner = ParseDecimal(r0.ScrewSpeedInner);
                    report.InitFeedRollRatioInner = ParseDecimal(r0.FeedRollRatioInner);
                    report.InitPressureInner = ParseDecimal(r0.PressureInner);

                    report.InitHeadTempOuter = ParseDecimal(r0.HeadTempOuter);
                    report.InitCylinder1TempOuter = ParseDecimal(r0.Cylinder1TempOuter);
                    report.InitCylinder2TempOuter = ParseDecimal(r0.Cylinder2TempOuter);
                    report.InitCylinder3TempOuter = ParseDecimal(r0.Cylinder3TempOuter);
                    report.InitScrewTempOuter = ParseDecimal(r0.ScrewTempOuter);
                    report.InitScrewSpeedOuter = ParseDecimal(r0.ScrewSpeedOuter);
                    report.InitFeedRollRatioOuter = ParseDecimal(r0.FeedRollRatioOuter);
                    report.InitPressureOuter = ParseDecimal(r0.PressureOuter);

                    report.InitSpiralSpeed = ParseDecimal(r0.SpiralSpeed);
                    report.InitSpiralPitchSetting = ParseDecimal(r0.SpiralPitchSetting);
                    report.InitHoseSpeed = ParseDecimal(r0.HoseSpeed);
                    report.InitConveyorRatio = ParseDecimal(r0.ConveyorRatio);
                    report.InitTakeupConveyorSpeed = ParseDecimal(r0.TakeupConveyorSpeed);
                    report.InitCoolConveyorSpeed = ParseDecimal(r0.CoolConveyorSpeed);
                    report.InitChillerWaterTemp = ParseDecimal(r0.ChillerWaterTemp);
                    report.InitCaterpillarGap = ParseDecimal(r0.CaterpillarGap);
                    report.InitUnsmoothSurface = ParseDecimal(r0.UnsmoothSurface);
                    report.InitPresetValue = ParseDecimal(r0.PresetValue);
                    report.InitControlValue = ParseDecimal(r0.ControlValue);
                }

                _context.ProductionReports.Add(report);
                await _context.SaveChangesAsync(); // Save first to get the ID

                // Batch insert Material Lots + Readings in a single round-trip
                if (dto.MaterialLots != null && dto.MaterialLots.Count > 0)
                {
                    var lots = dto.MaterialLots.Select(lotDto => new ProductionMaterialLot
                    {
                        ProductionReportId = report.Id,
                        LayerType    = lotDto.LayerType,
                        MaterialName = lotDto.MaterialName,
                        MaterialActual = lotDto.MaterialActual,
                        LotNumber    = lotDto.LotNumber,
                        SGValue      = lotDto.SGValue
                    });
                    _context.ProductionMaterialLots.AddRange(lots);
                }

                if (dto.Readings != null && dto.Readings.Count > 0)
                {
                    var readings = dto.Readings.Select(r => new ProductionReading
                    {
                        ProductionReportId  = report.Id,
                        ReadingTime         = DateTime.TryParse(r.ReadingTime, out var rt) ? rt : DateTime.Now,
                        RecordedBy          = r.RecordedBy ?? report.CreatedBy,

                        HeadTempInner       = int.TryParse(r.HeadTempInner,       out var hti) ? hti : (int?)null,
                        Cylinder1TempInner  = int.TryParse(r.Cylinder1TempInner,  out var c1i) ? c1i : (int?)null,
                        Cylinder2TempInner  = int.TryParse(r.Cylinder2TempInner,  out var c2i) ? c2i : (int?)null,
                        Cylinder3TempInner  = int.TryParse(r.Cylinder3TempInner,  out var c3i) ? c3i : (int?)null,
                        ScrewTempInner      = int.TryParse(r.ScrewTempInner,      out var sti) ? sti : (int?)null,
                        ScrewSpeedInner     = ParseDecimal(r.ScrewSpeedInner),
                        FeedRollRatioInner  = int.TryParse(r.FeedRollRatioInner,  out var fri) ? fri : (int?)null,
                        PressureInner       = ParseDecimal(r.PressureInner),

                        HeadTempOuter       = int.TryParse(r.HeadTempOuter,       out var hto) ? hto : (int?)null,
                        Cylinder1TempOuter  = int.TryParse(r.Cylinder1TempOuter,  out var c1o) ? c1o : (int?)null,
                        Cylinder2TempOuter  = int.TryParse(r.Cylinder2TempOuter,  out var c2o) ? c2o : (int?)null,
                        Cylinder3TempOuter  = int.TryParse(r.Cylinder3TempOuter,  out var c3o) ? c3o : (int?)null,
                        ScrewTempOuter      = int.TryParse(r.ScrewTempOuter,      out var sto) ? sto : (int?)null,
                        ScrewSpeedOuter     = ParseDecimal(r.ScrewSpeedOuter),
                        FeedRollRatioOuter  = int.TryParse(r.FeedRollRatioOuter,  out var fro) ? fro : (int?)null,
                        PressureOuter       = ParseDecimal(r.PressureOuter),

                        SpiralSpeed         = ParseDecimal(r.SpiralSpeed),
                        SpiralPitchSetting  = ParseDecimal(r.SpiralPitchSetting),
                        SpiralPitchDisplay  = ParseDecimal(r.SpiralPitchDisplay),
                        PresetValue         = ParseDecimal(r.PresetValue),
                        ControlValue        = ParseDecimal(r.ControlValue),
                        HoseSpeed           = ParseDecimal(r.HoseSpeed),
                        TakeupConveyorSpeed = ParseDecimal(r.TakeupConveyorSpeed),
                        CoolConveyorSpeed   = ParseDecimal(r.CoolConveyorSpeed),
                        ConveyorRatio       = r.ConveyorRatio,
                        UnsmoothSurface     = ParseDecimal(r.UnsmoothSurface),
                        ChillerWaterTemp    = ParseDecimal(r.ChillerWaterTemp),
                        CaterpillarGap      = ParseDecimal(r.CaterpillarGap)
                    });
                    _context.ProductionReadings.AddRange(readings);
                }

                // Single bulk save for lots + readings
                await _context.SaveChangesAsync();

                // Notify dashboard via SignalR (fire-and-forget, don't await to block response)
                // Final Save for all added collections
                await _context.SaveChangesAsync();
                
                // Notify dashboard via SignalR
                Console.WriteLine($">>> SIGNALR: Broadcasting 'ReceiveUpdate' from ProductionReport module (Save) at {DateTime.Now}");
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");
                return Json(new { success = true, id = report.Id, message = "Report archived successfully." });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "CRITICAL ERROR: " + ex.Message + (ex.InnerException != null ? " -> " + ex.InnerException.Message : "") });
            }
        }


        [HttpPost]
        public async Task<IActionResult> UpdateWithReadings([FromBody] ProductionReportSaveDto dto)
        {
            if (dto == null || dto.Id == 0) return Json(new { success = false, message = "Invalid data payload" });

            try
            {
                var report = await _context.ProductionReports
                    .Include(p => p.ProductionReadings)
                    .FirstOrDefaultAsync(m => m.Id == dto.Id);

                if (report == null) return Json(new { success = false, message = "Report not found" });

                // Update common fields
                report.DocumentNumber = dto.DocumentNumber ?? report.DocumentNumber;
                report.RevisionNumber = dto.RevisionNumber;
                if (DateTime.TryParse(dto.ProductionDate, out var dt)) report.ProductionDate = dt;
                report.Shift = ShiftHelper.NormalizeShiftLabel(dto.Shift);
                report.CustomerName = dto.CustomerName;
                report.HoseType = dto.HoseType;
                report.SpsId = SanitizeLegacySpsId(dto.SpsId);
                report.ItemCode = dto.ItemCode; // ← Pastikan ItemCode selalu tersimpan

                report.InnerMaterial = dto.InnerMaterial;
                report.InnerMaterialActual = dto.InnerMaterialActual;
                report.InnerMaterialLotNo = dto.InnerMaterialLotNo;
                report.OuterMaterial = dto.OuterMaterial;
                report.OuterMaterialActual = dto.OuterMaterialActual;
                report.OuterMaterialLotNo = dto.OuterMaterialLotNo;
                report.MiddleMaterial = dto.MiddleMaterial;         // ← Tambah Middle Material
                report.MiddleMaterialActual = dto.MiddleMaterialActual;
                report.MiddleMaterialLotNo = dto.MiddleMaterialLotNo;
                report.Yarn = dto.Yarn;
                report.YarnActual = dto.YarnActual;
                report.YarnLotNo = dto.YarnLotNo;

                report.DAI_Awal = dto.DAI_Awal;
                report.DAI_Akhir = dto.DAI_Akhir;
                report.DAC_Awal = dto.DAC_Awal;
                report.DAC_Akhir = dto.DAC_Akhir;
                report.DRAI_Awal = dto.DRAI_Awal;
                report.DRAI_Akhir = dto.DRAI_Akhir;
                report.DRAC_Awal = dto.DRAC_Awal;
                report.DRAC_Akhir = dto.DRAC_Akhir;

                report.NippleDieOK = dto.NippleDieOK; report.NippleDieInitial = dto.NippleDieInitial; report.NippleDieFinal = dto.NippleDieFinal;
                report.TubeDieOK = dto.TubeDieOK; report.TubeDieInitial = dto.TubeDieInitial; report.TubeDieFinal = dto.TubeDieFinal;
                report.MiddleDieOK = dto.MiddleDieOK; report.MiddleDieInitial = dto.MiddleDieInitial; report.MiddleDieFinal = dto.MiddleDieFinal;
                report.CoverDieOK = dto.CoverDieOK; report.CoverDieInitial = dto.CoverDieInitial; report.CoverDieFinal = dto.CoverDieFinal;
                report.SpacerDieOK = dto.SpacerDieOK; report.SpacerDieInitial = dto.SpacerDieInitial; report.SpacerDieFinal = dto.SpacerDieFinal;
                report.ToleranceDieOK = dto.ToleranceDieOK; report.ToleranceInitial = dto.ToleranceInitial; report.ToleranceFinal = dto.ToleranceFinal;

                report.MeshInner10Before = dto.MeshInner10Before; report.MeshInner40Before = dto.MeshInner40Before;
                report.MeshOuter10Before = dto.MeshOuter10Before; report.MeshOuter40Before = dto.MeshOuter40Before;
                report.MeshInnerCheck = dto.MeshInnerCheck; report.MeshOuterCheck = dto.MeshOuterCheck;
                report.EmbossMarkContent = dto.EmbossMarkContent;
                if (DateTime.TryParse(dto.EmbossMarkDate, out var ed)) report.EmbossMarkDate = ed;
                report.QcCond = dto.QcCond; report.QcSurf = dto.QcSurf; report.QcRes = dto.QcRes;

                // Sync Initial instrument parameters from the first reading if they are empty
                if (dto.Readings != null && dto.Readings.Count > 0)
                {
                    var r0 = dto.Readings[0];
                    report.InitHeadTempInner = decimal.TryParse(r0.HeadTempInner, out var it1) ? it1 : report.InitHeadTempInner;
                    report.InitCylinder1TempInner = decimal.TryParse(r0.Cylinder1TempInner, out var it2) ? it2 : report.InitCylinder1TempInner;
                    report.InitCylinder2TempInner = decimal.TryParse(r0.Cylinder2TempInner, out var it3) ? it3 : report.InitCylinder2TempInner;
                    report.InitCylinder3TempInner = decimal.TryParse(r0.Cylinder3TempInner, out var it4) ? it4 : report.InitCylinder3TempInner;
                    report.InitScrewTempInner = decimal.TryParse(r0.ScrewTempInner, out var it5) ? it5 : report.InitScrewTempInner;
                    report.InitScrewSpeedInner = decimal.TryParse(r0.ScrewSpeedInner, out var it6) ? it6 : report.InitScrewSpeedInner;
                    report.InitFeedRollRatioInner = decimal.TryParse(r0.FeedRollRatioInner, out var it7) ? it7 : report.InitFeedRollRatioInner;
                    report.InitPressureInner = decimal.TryParse(r0.PressureInner, out var it8) ? it8 : report.InitPressureInner;

                    report.InitHeadTempOuter = decimal.TryParse(r0.HeadTempOuter, out var ot1) ? ot1 : report.InitHeadTempOuter;
                    report.InitCylinder1TempOuter = decimal.TryParse(r0.Cylinder1TempOuter, out var ot2) ? ot2 : report.InitCylinder1TempOuter;
                    report.InitCylinder2TempOuter = decimal.TryParse(r0.Cylinder2TempOuter, out var ot3) ? ot3 : report.InitCylinder2TempOuter;
                    report.InitCylinder3TempOuter = decimal.TryParse(r0.Cylinder3TempOuter, out var ot4) ? ot4 : report.InitCylinder3TempOuter;
                    report.InitScrewTempOuter = decimal.TryParse(r0.ScrewTempOuter, out var ot5) ? ot5 : report.InitScrewTempOuter;
                    report.InitScrewSpeedOuter = decimal.TryParse(r0.ScrewSpeedOuter, out var ot6) ? ot6 : report.InitScrewSpeedOuter;
                    report.InitFeedRollRatioOuter = decimal.TryParse(r0.FeedRollRatioOuter, out var ot7) ? ot7 : report.InitFeedRollRatioOuter;
                    report.InitPressureOuter = decimal.TryParse(r0.PressureOuter, out var ot8) ? ot8 : report.InitPressureOuter;

                    report.InitSpiralSpeed = decimal.TryParse(r0.SpiralSpeed, out var m1) ? m1 : report.InitSpiralSpeed;
                    report.InitSpiralPitchSetting = decimal.TryParse(r0.SpiralPitchSetting, out var m2) ? m2 : report.InitSpiralPitchSetting;
                    report.InitHoseSpeed = decimal.TryParse(r0.HoseSpeed, out var m3) ? m3 : report.InitHoseSpeed;
                    report.InitConveyorRatio = decimal.TryParse(r0.ConveyorRatio, out var m4) ? m4 : report.InitConveyorRatio;
                    report.InitTakeupConveyorSpeed = decimal.TryParse(r0.TakeupConveyorSpeed, out var m5) ? m5 : report.InitTakeupConveyorSpeed;
                    report.InitCoolConveyorSpeed = decimal.TryParse(r0.CoolConveyorSpeed, out var m6) ? m6 : report.InitCoolConveyorSpeed;
                    report.InitChillerWaterTemp = decimal.TryParse(r0.ChillerWaterTemp, out var m7) ? m7 : report.InitChillerWaterTemp;
                    report.InitCaterpillarGap = decimal.TryParse(r0.CaterpillarGap, out var m8) ? m8 : report.InitCaterpillarGap;
                    report.InitUnsmoothSurface = decimal.TryParse(r0.UnsmoothSurface, out var m9) ? m9 : report.InitUnsmoothSurface;
                    report.InitPresetValue = decimal.TryParse(r0.PresetValue, out var m10) ? m10 : report.InitPresetValue;
                    report.InitControlValue = decimal.TryParse(r0.ControlValue, out var m11) ? m11 : report.InitControlValue;
                }

                // Update Readings: Remove old and add new (simplest for Edit)
                // Alternatively, match IDs if complexity is needed, but for 1-5 readings, clear & reload is safer
                _context.ProductionReadings.RemoveRange(report.ProductionReadings);

                if (dto.Readings != null)
                {
                    foreach (var r in dto.Readings)
                    {
                        var reading = new ProductionReading
                        {
                            ProductionReportId = report.Id,
                            ReadingTime = DateTime.TryParse(r.ReadingTime, out var rt) ? rt : DateTime.Now,
                            RecordedBy = r.RecordedBy ?? report.CreatedBy ?? "Operator",
                            
                            HeadTempInner = int.TryParse(r.HeadTempInner, out var hti) ? hti : (int?)null,
                            Cylinder1TempInner = int.TryParse(r.Cylinder1TempInner, out var c1i) ? c1i : (int?)null,
                            Cylinder2TempInner = int.TryParse(r.Cylinder2TempInner, out var c2i) ? c2i : (int?)null,
                            Cylinder3TempInner = int.TryParse(r.Cylinder3TempInner, out var c3i) ? c3i : (int?)null,
                            ScrewTempInner = int.TryParse(r.ScrewTempInner, out var sti) ? sti : (int?)null,
                            ScrewSpeedInner = ParseDecimal(r.ScrewSpeedInner),
                            FeedRollRatioInner = int.TryParse(r.FeedRollRatioInner, out var fri) ? fri : (int?)null,
                            PressureInner = ParseDecimal(r.PressureInner),

                            HeadTempOuter = int.TryParse(r.HeadTempOuter, out var hto) ? hto : (int?)null,
                            Cylinder1TempOuter = int.TryParse(r.Cylinder1TempOuter, out var c1o) ? c1o : (int?)null,
                            Cylinder2TempOuter = int.TryParse(r.Cylinder2TempOuter, out var c2o) ? c2o : (int?)null,
                            Cylinder3TempOuter = int.TryParse(r.Cylinder3TempOuter, out var c3o) ? c3o : (int?)null,
                            ScrewTempOuter = int.TryParse(r.ScrewTempOuter, out var sto) ? sto : (int?)null,
                            ScrewSpeedOuter = ParseDecimal(r.ScrewSpeedOuter),
                            FeedRollRatioOuter = int.TryParse(r.FeedRollRatioOuter, out var fro) ? fro : (int?)null,
                            PressureOuter = ParseDecimal(r.PressureOuter),

                            SpiralSpeed = ParseDecimal(r.SpiralSpeed),
                            SpiralPitchSetting = ParseDecimal(r.SpiralPitchSetting),
                            SpiralPitchDisplay = ParseDecimal(r.SpiralPitchDisplay),
                            PresetValue = ParseDecimal(r.PresetValue),
                            ControlValue = ParseDecimal(r.ControlValue),
                            HoseSpeed = ParseDecimal(r.HoseSpeed),
                            TakeupConveyorSpeed = ParseDecimal(r.TakeupConveyorSpeed),
                            CoolConveyorSpeed = ParseDecimal(r.CoolConveyorSpeed),
                            ConveyorRatio = r.ConveyorRatio,
                            UnsmoothSurface = ParseDecimal(r.UnsmoothSurface),
                            ChillerWaterTemp = ParseDecimal(r.ChillerWaterTemp),
                            CaterpillarGap = ParseDecimal(r.CaterpillarGap)
                        };
                        _context.ProductionReadings.Add(reading);
                    }
                }

                await _context.SaveChangesAsync();
                
                // Notify via SignalR
                // Notify via SignalR
                Console.WriteLine($">>> SIGNALR: Broadcasting 'ReceiveUpdate' from ProductionReport module (Update) at {DateTime.Now}");
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        // Helper to handle messy inputs like "15.5 mm" or "100 rpm"
        private decimal? ParseDecimal(string? input)
        {
            if (string.IsNullOrWhiteSpace(input) || input == "---") return null;
            try {
                // Normalize: Replace comma with dot first
                string normalized = input.Replace(',', '.');
                // Keep only numbers, one dot, and a hyphen
                string cleaned = System.Text.RegularExpressions.Regex.Replace(normalized, @"[^-0-9.]", "");
                
                if (decimal.TryParse(cleaned, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out decimal result))
                    return result;
            } catch { }
            return null;
        }

        public async Task<IActionResult> Details(int? id)
        {
            if (id == null) return NotFound();

            var report = await _context.ProductionReports
                .Include(p => p.ProductionReadings)
                .Include(p => p.MaterialLots)
                .FirstOrDefaultAsync(m => m.Id == id);

            if (report == null) return NotFound();

            var sps = await FindSpsNoDocForReportAsync(report);
            ViewBag.Sps = sps;
            ViewBag.Masterlist = sps;

            return View(report);
        }

        // GET: ProductionReport/Create (The "NOW I'M PRODUCE" Form - Gambar 1)
        public async Task<IActionResult> Create()
        {
            

            var hoseTypes = await _context.SpsNoDocs
                .Where(m => !string.IsNullOrEmpty(m.HoseType))
                .Select(m => m.HoseType)
                .Distinct()
                .OrderBy(h => h)
                .ToListAsync();

            ViewBag.HoseTypes = new SelectList(hoseTypes);
            ViewBag.StandardParameterSettings = new SelectList(
                await _context.SpsNoDocs.Select(s => new { Id = s.DocumentNumber, Display = "[" + s.DocumentNumber + "] - " + s.HoseType }).ToListAsync(), 
                "Id", 
                "Display"
            );

            return View(new ProductionReport 
            { 
                ProductionDate = DateTime.Today,
                Status = "NOW PRODUCING",
                Yarn = "---",
                Shift = ShiftHelper.GetCurrentShift(await _context.ShiftMasters.ToListAsync()),
                CreatedBy = HttpContext.Session.GetString("UserName") ?? "Operator"
            });
        }

        // POST: ProductionReport/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(ProductionReport report)
        {
            // Remove validation for fields that are not in Gambar 1
            ModelState.Remove("CreatedBy");
            
            if (ModelState.IsValid)
            {
                try
                {
                    if (report.ProductionDate == default)
                        report.ProductionDate = DateTime.Today;

                    report.CreatedBy = report.CreatedBy ?? "Operator 1";
                    
                    if (string.IsNullOrEmpty(report.Shift))
                        report.Shift = ShiftHelper.GetCurrentShift(await _context.ShiftMasters.ToListAsync());
                    else
                        report.Shift = ShiftHelper.NormalizeShiftLabel(report.Shift);
                    report.CustomerName = report.CustomerName ?? "-";
                    report.CreatedDate = DateTime.Now;
                    report.Status = "NOW PRODUCING";

                    // Jendela 4 jam per mesin: save pertama buka jendela, save berikutnya ikut jendela aktif
                    if (report.ProductionStartTime == null || report.ProductionStartTime == DateTime.MinValue
                        || report.ProductionEndTime == null || report.ProductionEndTime == DateTime.MinValue)
                    {
                        var (productionStart, productionEnd) = await GetProductionWindowAsync(report.MachineName);
                        report.ProductionStartTime = productionStart;
                        report.ProductionEndTime = productionEnd;
                    }

                    // Set default values for NOT NULL columns not present in the form
                    report.InnerMaterialActual = report.InnerMaterialActual ?? report.InnerMaterial ?? "-";
                    report.OuterMaterialActual = report.OuterMaterialActual ?? report.OuterMaterial ?? "-";
                    report.YarnActual = report.YarnActual ?? report.Yarn ?? "-";

                    _context.Add(report);
                    await _context.SaveChangesAsync();
                    
                    // Notify via SignalR
                    await _hubContext.Clients.All.SendAsync("ReceiveUpdate");
                    
                    TempData["SuccessMessage"] = "Produksi Dimulai! Silakan monitor di Dashboard Parameter History.";
                    return RedirectToAction(nameof(ParameterHistory));
                }
                catch (Exception ex)
                {
                    var innerMsg = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                    if (ex.InnerException?.InnerException != null) 
                        innerMsg = ex.InnerException.InnerException.Message;

                    Console.WriteLine($"Error saving production report: {innerMsg}");
                    ModelState.AddModelError("", "Gagal Simpan ke Database: " + innerMsg);
                }
            }
            return View(report);
        }

        // GET: ProductionReport/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null) return NotFound();

            var report = await _context.ProductionReports
                .Include(p => p.ProductionReadings)
                .Include(p => p.MaterialLots)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (report == null) return NotFound();

            report.Shift = ShiftHelper.NormalizeShiftLabel(report.Shift);

            ViewBag.StandardParameterSettings = new SelectList(
                await _context.SpsNoDocs
                    .Select(s => new { Id = s.DocumentNumber, Display = "[" + s.DocumentNumber + "] - " + s.HoseType })
                    .ToListAsync(),
                "Id", "Display", report.DocumentNumber
            );

            // Fetch specific SPS record by DocumentNumber
            ViewBag.Sps = await _context.SpsNoDocs
                .OrderByDescending(s => s.DocumentNumber)
                .FirstOrDefaultAsync(s => s.DocumentNumber == report.ItemCode || s.DocumentNumber == report.DocumentNumber);

            return View(report);
        }

        // POST: ProductionReport/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, ProductionReport report)
        {
            if (id != report.Id) return NotFound();

            ModelState.Remove("CreatedBy");
            ModelState.Remove("CreatedDate");
            ModelState.Remove("Status");

            if (ModelState.IsValid)
            {
                try
                {
                    // Load existing from DB to preserve immutable fields
                    var existing = await _context.ProductionReports.FindAsync(id);
                    if (existing == null) return NotFound();

                    // === EDITABLE FIELDS ONLY ===
                    // Header
                    existing.DocumentNumber = report.DocumentNumber;
                    existing.ProductionDate = report.ProductionDate;
                    existing.RevisionNumber = report.RevisionNumber;
                    existing.Shift = ShiftHelper.NormalizeShiftLabel(report.Shift);
                    existing.CustomerName = report.CustomerName;
                    existing.HoseType = report.HoseType;
                    existing.Dimension = report.Dimension;
                    existing.ItemCode = report.ItemCode;
                    existing.SpsId = report.SpsId;

                    // Standards (from SPS)
                    existing.InnerMaterial = report.InnerMaterial;
                    existing.OuterMaterial = report.OuterMaterial;
                    existing.Yarn = report.Yarn;

                    // Material Actual + Lot
                    existing.InnerMaterialActual = report.InnerMaterialActual;
                    existing.InnerMaterialLotNo = report.InnerMaterialLotNo;
                    existing.InnerMaterialSG = report.InnerMaterialSG;
                    existing.OuterMaterialActual = report.OuterMaterialActual;
                    existing.OuterMaterialLotNo = report.OuterMaterialLotNo;
                    existing.OuterMaterialSG = report.OuterMaterialSG;
                    existing.YarnActual = report.YarnActual;
                    existing.YarnLotNo = report.YarnLotNo;

                    // Die Verification
                    existing.NippleDieOK = report.NippleDieOK;
                    existing.NippleDieInitial = report.NippleDieInitial;
                    existing.NippleDieFinal = report.NippleDieFinal;
                    existing.TubeDieOK = report.TubeDieOK;
                    existing.TubeDieInitial = report.TubeDieInitial;
                    existing.TubeDieFinal = report.TubeDieFinal;
                    existing.MiddleDieOK = report.MiddleDieOK;
                    existing.MiddleDieInitial = report.MiddleDieInitial;
                    existing.MiddleDieFinal = report.MiddleDieFinal;
                    existing.CoverDieOK = report.CoverDieOK;
                    existing.CoverDieInitial = report.CoverDieInitial;
                    existing.CoverDieFinal = report.CoverDieFinal;
                    existing.SpacerDieOK = report.SpacerDieOK;
                    existing.SpacerDieInitial = report.SpacerDieInitial;
                    existing.SpacerDieFinal = report.SpacerDieFinal;
                    existing.ToleranceDieOK = report.ToleranceDieOK;
                    existing.ToleranceInitial = report.ToleranceInitial;
                    existing.ToleranceFinal = report.ToleranceFinal;

                    // Mesh & Emboss
                    existing.MeshInner10Before = report.MeshInner10Before;
                    existing.MeshInner40Before = report.MeshInner40Before;
                    existing.MeshOuter10Before = report.MeshOuter10Before;
                    existing.MeshOuter40Before = report.MeshOuter40Before;
                    existing.MeshInnerCheck = report.MeshInnerCheck;
                    existing.MeshOuterCheck = report.MeshOuterCheck;
                    existing.EmbossMarkContent = report.EmbossMarkContent;
                    existing.EmbossMarkDate = report.EmbossMarkDate;

                    // QC
                    existing.QcCond = report.QcCond;
                    existing.QcSurf = report.QcSurf;
                    existing.QcRes = report.QcRes;

                    // Waste / Dandori
                    existing.DAI_Awal = report.DAI_Awal;
                    existing.DAI_Akhir = report.DAI_Akhir;
                    existing.DAC_Awal = report.DAC_Awal;
                    existing.DAC_Akhir = report.DAC_Akhir;
                    existing.DRAI_Awal = report.DRAI_Awal;
                    existing.DRAI_Akhir = report.DRAI_Akhir;
                    existing.DRAC_Awal = report.DRAC_Awal;
                    existing.DRAC_Akhir = report.DRAC_Akhir;

                    // Production Summary
                    existing.QtyTarget = report.QtyTarget;
                    existing.QtyOk = report.QtyOk;
                    existing.NgDimension = report.NgDimension;
                    existing.NgVisual = report.NgVisual;
                    existing.VinCode = report.VinCode;
                    existing.StandardLength = report.StandardLength;
                    existing.ActualLength = report.ActualLength;
                    existing.Remark = report.Remark;

                    // Operator (allow update)
                    if (!string.IsNullOrWhiteSpace(report.CreatedBy))
                        existing.CreatedBy = report.CreatedBy;

                    await _context.SaveChangesAsync();

                    // Notify via SignalR for real-time dashboard updates
                    await _hubContext.Clients.All.SendAsync("ReceiveUpdate");

                    TempData["SuccessMessage"] = "Data berhasil diupdate!";
                    return RedirectToAction(nameof(ParameterHistory));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!ProductionReportExists(report.Id)) return NotFound();
                    else throw;
                }
            }

            ViewBag.StandardParameterSettings = new SelectList(
                await _context.SpsNoDocs
                    .Select(s => new { Id = s.DocumentNumber, Display = "[" + s.DocumentNumber + "] - " + s.HoseType })
                    .ToListAsync(),
                "Id", "Display", report.DocumentNumber
            );
            return View(report);
        }

        // GET: ProductionReport/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null) return NotFound();

            var report = await _context.ProductionReports
                .FirstOrDefaultAsync(m => m.Id == id);
            if (report == null) return NotFound();

            return View(report);
        }

        // POST: ProductionReport/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var report = await _context.ProductionReports.FindAsync(id);
            if (report != null)
            {
                _context.ProductionReports.Remove(report);
                await _context.SaveChangesAsync();

                // Notify via SignalR for real-time dashboard updates
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");

                TempData["SuccessMessage"] = "Laporan berhasil dihapus!";
            }
            
            return RedirectToAction(nameof(ParameterHistory));
        }

        private bool ProductionReportExists(int id)
        {
            return _context.ProductionReports.Any(e => e.Id == id);
        }

        [HttpPost]
        public async Task<IActionResult> FinishProduction(int id, int qtyOk, int ngDim, int ngVis, 
            decimal? WasteInnerAwal, decimal? WasteInnerAkhir, decimal? WasteCoverAwal, decimal? WasteCoverAkhir,
            string? DAI_Awal, string? DAI_Akhir, string? DAC_Awal, string? DAC_Akhir,
            string? DRAI_Awal, string? DRAI_Akhir, string? DRAC_Awal, string? DRAC_Akhir)
        {
            var report = await _context.ProductionReports.FindAsync(id);
            if (report != null)
            {
                report.QtyOk = qtyOk;
                report.NgDimension = ngDim;
                report.NgVisual = ngVis;
                report.WasteInnerAwal = WasteInnerAwal;
                report.WasteInnerAkhir = WasteInnerAkhir;
                report.WasteCoverAwal = WasteCoverAwal;
                report.WasteCoverAkhir = WasteCoverAkhir;

                report.DAI_Awal = DAI_Awal;
                report.DAI_Akhir = DAI_Akhir;
                report.DAC_Awal = DAC_Awal;
                report.DAC_Akhir = DAC_Akhir;
                report.DRAI_Awal = DRAI_Awal;
                report.DRAI_Akhir = DRAI_Akhir;
                report.DRAC_Awal = DRAC_Awal;
                report.DRAC_Akhir = DRAC_Akhir;

                report.Status = "COMPLETED";
                report.ProductionEndTime = DateTime.Now;
                await _context.SaveChangesAsync();

                // Notify via SignalR for real-time dashboard updates
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");

                TempData["SuccessMessage"] = "Produksi Selesai!";
            }
            return RedirectToAction(nameof(Index));
        }

        [HttpGet]
        public async Task<IActionResult> GetMasterByHoseType(string hoseType)
        {
            var master = await _context.SpsNoDocs
                .FirstOrDefaultAsync(m => m.HoseType == hoseType);
            return Json(master);
        }

        [HttpGet]
        public async Task<IActionResult> GetMasterByItem(string itemCode)
        {
            var master = await _context.SpsNoDocs
                .FirstOrDefaultAsync(m => m.DocumentNumber == itemCode);
            return Json(master);
        }

        [HttpPost]
        public async Task<IActionResult> UpdateStandards(int reportId, string itemCode)
        {
            var report = await _context.ProductionReports.FindAsync(reportId);
            var master = await _context.SpsNoDocs
                .FirstOrDefaultAsync(m => m.DocumentNumber == itemCode);

            if (report != null && master != null)
            {
                report.HoseType = master.HoseType ?? "";
                report.Dimension = master.Dimensi ?? "";
                report.InnerMaterial = master.InnerTube ?? master.Material ?? "";
                report.OuterMaterial = master.OuterCover ?? "";
                
                // Logic to update or create StandardParameterSetting can be added here
                
                await _context.SaveChangesAsync();
                return Json(new { success = true });
            }
            return Json(new { success = false });
        }

        [HttpPost]
        public async Task<IActionResult> VerifyApproval(int id, string role, string pin)
        {
            var report = await _context.ProductionReports.FindAsync(id);
            if (report == null) return Json(new { success = false, message = "Report not found" });

            // SIMPLE PIN LOGIC: 4321 for Leader (Diperiksa), 1234 for Foreman (Disetujui)
            // In real app, check against StaffMaster/Users table
            bool isValid = false;
            string staffName = "";

            if (role == "LEADER" && pin == "4321") { isValid = true; staffName = "LEADER PRODUCTION"; }
            else if (role == "FOREMAN" && pin == "1234") { isValid = true; staffName = "FOREMAN EXTRUDER"; }

            if (isValid)
            {
                if (role == "LEADER")
                {
                    report.CheckedBy = staffName;
                    report.CheckedDate = DateTime.Now;
                    report.CheckedBySignature = staffName.ToUpper();
                }
                else
                {
                    report.ApprovedBy = staffName;
                    report.ApprovedDate = DateTime.Now;
                    report.ApprovedBySignature = staffName.ToUpper();
                    report.Status = "COMPLETED";
                }

                await _context.SaveChangesAsync();
                return Json(new { success = true, name = staffName, date = DateTime.Now.ToString("dd/MM/yyyy HH:mm") });
            }

            return Json(new { success = false, message = "PIN SALAH!" });
        }

        [HttpPost]
        public async Task<IActionResult> AddReading(ProductionReading reading)
        {
            if (reading.ProductionReportId != 0)
            {
                reading.ReadingTime = DateTime.Now;
                reading.RecordedBy = reading.RecordedBy ?? "Operator 1";
                _context.ProductionReadings.Add(reading);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Details), new { id = reading.ProductionReportId, portal = "parameter" });
            }
            return BadRequest();
        }

        [HttpPost]
        public async Task<IActionResult> SaveAppData([FromBody] DimensionReportAppData data)
        {
            try
            {
                ProductionReport? report;
                if (data.ReportId == 0)
                {
                    // Create new report from Scanned App
                    report = new ProductionReport
                    {
                        DocumentNumber = data.DocumentNumber ?? "-",
                        HoseType = data.HoseType ?? "-",
                        Dimension = data.DimensionDisplay ?? "-",
                        CustomerName = data.CustomerName ?? "-",
                        Yarn = data.Yarn ?? "-",
                        ProductionDate = DateTime.Today,
                        CreatedDate = DateTime.Now,
                        Status = "NOW PRODUCING",
                        InnerMaterialActual = "-",
                        OuterMaterialActual = "-",
                        YarnActual = "-",
                        CreatedBy = data.CreatedBy ?? "Operator",
                        Shift = ShiftHelper.NormalizeShiftLabel(data.Shift ?? ShiftHelper.GetCurrentShift(await _context.ShiftMasters.ToListAsync())),
                        ItemCode = data.ItemCode,
                        SpsId = data.SpsId
                    };
                    _context.ProductionReports.Add(report);
                }
                else
                {
                    report = await _context.ProductionReports.FindAsync(data.ReportId);
                    if (report == null) return Json(new { success = false, message = "Report not found" });
                }

                // Update common dimension readings & production info
                report.VinCode = data.VinCode;
                report.ActualLength = data.ActualLength;
                report.QtyOk = data.QtyOk ?? 0;
                report.NgDimension = data.NgDimension ?? 0;
                report.NgVisual = data.NgVisual ?? 0;
                report.Remark = data.Remark ?? "";

                await _context.SaveChangesAsync();

                // Notify via SignalR for real-time dashboard updates
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");

                return Json(new { success = true, reportId = report.Id, message = "Data saved successfully" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        public async Task<IActionResult> SubmitAppReport([FromBody] DimensionReportAppData data)
        {
            try
            {
                var report = await _context.ProductionReports.FindAsync(data.ReportId);
                if (report == null) return Json(new { success = false, message = "Report not found" });

                // Update all data
                report.ActualLength = data.ActualLength;
                report.QtyOk = data.QtyOk ?? 0;
                report.NgDimension = data.NgDimension ?? 0;
                report.NgVisual = data.NgVisual ?? 0;
                report.Remark = data.Remark ?? "";
                
                // Update status
                report.Status = "SUBMITTED";
                report.CheckedBy = "QC Inspector";
                report.ApprovedBy = "Production Manager";

                await _context.SaveChangesAsync();

                // Notify via SignalR for real-time dashboard updates
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");

                return Json(new { success = true, message = "Report submitted successfully" });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public async Task<IActionResult> ExportAppData(int id)
        {
            var report = await _context.ProductionReports
                .FirstOrDefaultAsync(m => m.Id == id);

            if (report == null) return NotFound();

            // Create CSV export
            var csv = $"Document Number,{report.DocumentNumber}\n" +
                     $"Date,{report.ProductionDate:yyyy-MM-dd}\n" +
                     $"Shift,{report.Shift}\n" +
                     $"Customer,{report.CustomerName}\n" +
                     $"Hose Type,{report.HoseType}\n" +
                     $"Dimension,{report.Dimension}\n" +
                     $"Actual Length,{report.ActualLength}\n" +
                     $"Quantity OK,{report.QtyOk}\n" +
                     $"NG Dimension,{report.NgDimension}\n" +
                     $"NG Visual,{report.NgVisual}\n" +
                     $"Remark,{report.Remark}\n";

            var bytes = System.Text.Encoding.UTF8.GetBytes(csv);
            return File(bytes, "text/csv", $"ProductionReport_{report.DocumentNumber}_{DateTime.Now:yyyyMMdd}.csv");
        }


        [HttpPost]
        public async Task<IActionResult> UpdateRecap(int reportId, string vinCode, string actualLength, int qtyOk, int ngDim, int ngVis, string remark)
        {
            var report = await _context.ProductionReports.FindAsync(reportId);
            if (report != null)
            {
                if (!string.IsNullOrEmpty(vinCode)) report.VinCode = vinCode;
                if (!string.IsNullOrEmpty(actualLength)) report.ActualLength = actualLength;
                
                report.QtyOk = qtyOk;
                report.NgDimension = ngDim;
                report.NgVisual = ngVis;
                report.Remark = remark;
                await _context.SaveChangesAsync();

                // Notify via SignalR for real-time dashboard updates
                await _hubContext.Clients.All.SendAsync("ReceiveUpdate");

                return Json(new { success = true });
            }
            return Json(new { success = false });
        }

        // ═══════════════════════════════════════════════════════════════
        //  CHART ANALYSIS — Page + API
        // ═══════════════════════════════════════════════════════════════

        // GET: ProductionReport/ChartAnalysis
        public IActionResult ChartAnalysis()
        {
            

            return View();
        }

        // GET: ProductionReport/GetChartData?documentNumber=XXX
        [HttpGet]
        public async Task<IActionResult> GetChartData(string documentNumber, string? startDate = null, string? endDate = null, string? shift = null, string? pic = null, string? itemCode = null)
        {
            if (string.IsNullOrWhiteSpace(documentNumber))
                return BadRequest(new { success = false, message = "Document number is required." });

            var docNum = documentNumber.Trim();

            // 1. Find ProductionReport
            var report = await _context.ProductionReports
                .Include(r => r.ProductionReadings)
                .FirstOrDefaultAsync(r => r.DocumentNumber == docNum);

            if (report == null)
                return Json(new { success = false, message = $"Production Report '{docNum}' tidak ditemukan." });

            Console.WriteLine($"[GetChartData] === START ===");
            Console.WriteLine($"[GetChartData] Document: {report.DocumentNumber}");
            Console.WriteLine($"[GetChartData] Report VinCode: {report.VinCode ?? "NULL"}");
            Console.WriteLine($"[GetChartData] Report ItemCode: {report.ItemCode ?? "NULL"}");
            Console.WriteLine($"[GetChartData] Requested ItemCode: {itemCode ?? "NULL"}");
            Console.WriteLine($"[GetChartData] Report HoseType: {report.HoseType ?? "NULL"}");

            // Use requested itemCode from frontend if report.ItemCode is NULL (common in simulator)
            var effectiveItemCode = (itemCode ?? report.ItemCode ?? "").Trim().ToUpper();
            var effectiveVinCode = (report.VinCode ?? "").Trim().ToUpper();

            // 2. Match SPS by VinCode/ItemCode FIRST (more accurate than linked SpsId)
            // PRIORITIZE: Master SPS (Edit) / SpsMasters
            SpsNoDoc? spsNew = null;
            SpsNoDoc? spsOld = null;
            var hasSps = false;
            string spsMatchMethod = "NONE";

            // Priority 1: Try matching by VinCode in SpsMasters
            if (!string.IsNullOrEmpty(effectiveVinCode))
            {
                spsNew = await _context.SpsNoDocs
                    .Where(m => m.DocumentNumber != null && m.DocumentNumber.ToUpper() == effectiveVinCode)
                    .OrderByDescending(m => m.DocumentNumber)
                    .FirstOrDefaultAsync();
                
                if (spsNew != null)
                {
                    spsMatchMethod = $"SpsMasters by VinCode '{effectiveVinCode}'";
                    hasSps = true;
                }
            }

            // Priority 2: Try matching by ItemCode in SpsMasters
            if (!hasSps && !string.IsNullOrEmpty(effectiveItemCode))
            {
                spsNew = await _context.SpsNoDocs
                    .Where(m => m.DocumentNumber != null && m.DocumentNumber.ToUpper() == effectiveItemCode)
                    .OrderByDescending(m => m.DocumentNumber)
                    .FirstOrDefaultAsync();
                
                if (spsNew != null)
                {
                    spsMatchMethod = $"SpsMasters by ItemCode '{effectiveItemCode}'";
                    hasSps = true;
                }
            }

            // Priority 3: Fallback to MasterlistSpsDoubleLayers if not found in SpsMasters
            if (!hasSps && !string.IsNullOrEmpty(effectiveVinCode))
            {
                spsOld = await _context.SpsNoDocs
                    .Where(m => m.DocumentNumber != null && m.DocumentNumber.ToUpper() == effectiveVinCode)
                    .OrderByDescending(m => m.DocumentNumber)
                    .FirstOrDefaultAsync();
                
                if (spsOld != null)
                {
                    spsMatchMethod = $"MasterlistSpsDoubleLayers by VinCode '{effectiveVinCode}'";
                    hasSps = true;
                }
            }

            if (!hasSps && !string.IsNullOrEmpty(effectiveItemCode))
            {
                spsOld = await _context.SpsNoDocs
                    .Where(m => m.DocumentNumber != null && m.DocumentNumber.ToUpper() == effectiveItemCode)
                    .OrderByDescending(m => m.DocumentNumber)
                    .FirstOrDefaultAsync();
                
                if (spsOld != null)
                {
                    spsMatchMethod = $"MasterlistSpsDoubleLayers by ItemCode '{effectiveItemCode}'";
                    hasSps = true;
                }
            }

            // Priority 4: Try DocumentNumber match
            if (!hasSps)
            {
                spsNew = await _context.SpsNoDocs.FirstOrDefaultAsync(m => m.DocumentNumber == report.DocumentNumber);
                if (spsNew != null) { spsMatchMethod = "SpsMasters by DocumentNumber"; hasSps = true; }
                else {
                    spsOld = await _context.SpsNoDocs.FirstOrDefaultAsync(m => m.DocumentNumber == report.DocumentNumber);
                    if (spsOld != null) { spsMatchMethod = "MasterlistSpsDoubleLayers by DocumentNumber"; hasSps = true; }
                }
            }

            if (!hasSps) Console.WriteLine($"[GetChartData] ❌ NO SPS FOUND!");
            else Console.WriteLine($"[GetChartData] ✅ SPS Found: {spsMatchMethod}");

            // 3. Load sensor data linked to this report
            var sensorData = await _context.SensorIngestLogs
                .Where(s => s.ProductionReportId == report.Id && s.Status == "OK")
                .OrderBy(s => s.SensorTimestamp)
                .Select(s => new {
                    s.MetricType,
                    value = s.MetricValue,
                    time = s.SensorTimestamp.ToString("HH:mm:ss"),
                    s.Unit,
                    s.Quality
                })
                .Take(500)
                .ToListAsync();

            // 4. Build readings time-series
            var readings = report.ProductionReadings
                .OrderBy(r => r.ReadingTime)
                .Select(r => new {
                    readingTime = r.ReadingTime.ToString("yyyy-MM-dd HH:mm"),
                    timeLabel = r.ReadingTime.ToString("HH:mm"),
                    r.HeadTempInner, r.HeadTempOuter,
                    r.Cylinder1TempInner, r.Cylinder1TempOuter,
                    r.Cylinder2TempInner, r.Cylinder2TempOuter,
                    r.Cylinder3TempInner, r.Cylinder3TempOuter,
                    r.ScrewTempInner, r.ScrewTempOuter,
                    r.ScrewSpeedInner, r.ScrewSpeedOuter,
                    r.FeedRollRatioInner, r.FeedRollRatioOuter,
                    r.PressureInner, r.PressureOuter,
                    r.HoseSpeed, r.SpiralSpeed,
                    r.SpiralPitchSetting, r.SpiralPitchDisplay,
                    r.ControlValue, r.PresetValue,
                    r.ChillerWaterTemp, r.CaterpillarGap,
                    r.TakeupConveyorSpeed, r.CoolConveyorSpeed,
                    r.InnerDiameter,
                    r.TotalThicknessX, r.TotalThicknessY,
                    r.SpiralPitch
                }).ToList();

            // 4b. Build dimension series from Dimension History when available
            var dimensionReportQuery = _context.DimensionReports
                .Include(d => d.Measurements)
                .AsQueryable();

            // Apply date filters
            if (!string.IsNullOrWhiteSpace(startDate) && DateTime.TryParse(startDate, out var dimStartDate))
            {
                dimensionReportQuery = dimensionReportQuery.Where(d => d.CreatedDate >= dimStartDate.Date);
            }

            if (!string.IsNullOrWhiteSpace(endDate) && DateTime.TryParse(endDate, out var dimEndDate))
            {
                dimensionReportQuery = dimensionReportQuery.Where(d => d.CreatedDate <= dimEndDate.Date.AddDays(1).AddTicks(-1));
            }

            // Apply shift filter
            if (!string.IsNullOrWhiteSpace(shift) && !shift.Equals("SHIFT ALL", StringComparison.OrdinalIgnoreCase))
            {
                var shiftNorm = shift.Trim().ToUpper();
                dimensionReportQuery = dimensionReportQuery.Where(d => d.Shift != null && d.Shift.ToUpper().Contains(shiftNorm));
            }

            // Apply PIC/Operator filter
            if (!string.IsNullOrWhiteSpace(pic))
            {
                var picNorm = pic.Trim().ToUpper();
                dimensionReportQuery = dimensionReportQuery.Where(d => d.CreatedBy != null && d.CreatedBy.ToUpper().Contains(picNorm));
            }

            List<DimensionReport> dimensionReports = new();

            // Priority 1: Match by VinCode or effective ItemCode (most reliable for item-based search)
            if (!string.IsNullOrWhiteSpace(effectiveVinCode))
            {
                var rows = await dimensionReportQuery
                    .Where(d => d.VinCode != null && d.VinCode.ToUpper() == effectiveVinCode)
                    .OrderByDescending(d => d.CreatedDate)
                    .Take(30)
                    .ToListAsync();
                dimensionReports = rows.OrderBy(d => d.CreatedDate).ToList();
            }

            if (!dimensionReports.Any() && !string.IsNullOrWhiteSpace(effectiveItemCode))
            {
                var rows = await dimensionReportQuery
                    .Where(d => d.ItemCode != null && d.ItemCode.ToUpper() == effectiveItemCode)
                    .OrderByDescending(d => d.CreatedDate)
                    .Take(30)
                    .ToListAsync();
                dimensionReports = rows.OrderBy(d => d.CreatedDate).ToList();
            }

            // Priority 3: Match by DocumentNumber
            if (!dimensionReports.Any())
            {
                var rows = await dimensionReportQuery
                    .Where(d => d.DocumentNumber == report.DocumentNumber)
                    .OrderByDescending(d => d.CreatedDate)
                    .Take(30)
                    .ToListAsync();
                dimensionReports = rows.OrderBy(d => d.CreatedDate).ToList();
            }

            // Priority 4: Match by MachineName as fallback
            if (!dimensionReports.Any() && !string.IsNullOrWhiteSpace(report.MachineName))
            {
                var machine = report.MachineName.Trim().ToUpper();
                var rows = await dimensionReportQuery
                    .Where(d => d.MachineName != null && d.MachineName.ToUpper() == machine)
                    .OrderByDescending(d => d.CreatedDate)
                    .Take(30)
                    .ToListAsync();
                dimensionReports = rows.OrderBy(d => d.CreatedDate).ToList();
            }

            object? dimensionSeries = null;

            Console.WriteLine($"[GetChartData] Total reports loaded: {dimensionReports.Count}");

            if (dimensionReports.Any())
            {
                // Group measurements with their parent report's production date to include in labels
                List<object> BuildDimensionPoints(Func<DimensionMeasurement, bool> predicate)
                {
                    var points = new List<object>();
                    foreach (var dReport in dimensionReports)
                    {
                        if (dReport.Measurements == null) continue;
                        
                        var matches = dReport.Measurements.Where(predicate);
                        foreach (var m in matches)
                        {
                            foreach (var reading in ExpandDimensionReadings(m, dReport.CreatedDate))
                            {
                                points.Add(new
                                {
                                    label = reading.label,
                                    value = reading.value,
                                    pointName = m.PointName,
                                    recordedTime = m.RecordedTime // Keep for sorting
                                });
                            }
                        }
                    }

                    // Sort all points chronologically to ensure the chart line flows correctly across days
                    return points.OrderBy(p => ((dynamic)p).recordedTime).ToList();
                }

                dimensionSeries = new
                {
                    reportCount = dimensionReports.Count,
                    source = "DimensionHistory",
                    innerDiameter = BuildDimensionPoints(m => ContainsAny(m.PointName, "inner diameter", "inner tube diameter")),
                    innerThicknessX = BuildDimensionPoints(m => ContainsAny(m.PointName, "inner thickness (x)", "inner thickness x", "inner tube thickness x")),
                    innerThicknessY = BuildDimensionPoints(m => ContainsAny(m.PointName, "inner thickness (y)", "inner thickness y", "inner tube thickness y")),
                    totalThicknessX = BuildDimensionPoints(m => ContainsAny(m.PointName, "total thickness (x)", "outer cover total thickness", "total thickness x")),
                    totalThicknessY = BuildDimensionPoints(m => ContainsAny(m.PointName, "total thickness (y)", "total thickness y"))
                };
            }

            // Prioritize spsNew (Master SPS Edit) over spsOld (Masterlist)
            object? dimensionStandard = null;
            object? spsPayload = null;

            if (spsNew != null)
            {
                Console.WriteLine($"[GetChartData] Using SpsMasters (EDIT) for dimension standards");

                // === INNER DIAMETER — 3-level fallback (mirrors legacy Masterlist logic) ===
                decimal? InnerMin_v, InnerMax_v, InnerTarget_v;

                // Level 1: Quality Matrix explicit strings (InnerMin / InnerMax / InnerTarget)
                (InnerMin_v, InnerMax_v, InnerTarget_v) = ParseRange(spsNew.InnerMin ?? spsNew.InnerTarget);
                if (InnerMin_v == null && !string.IsNullOrEmpty(spsNew.InnerMax))
                    (_, InnerMax_v, _) = ParseRange(spsNew.InnerMax);

                // Level 2: InnerTarget string + tolerance extracted from ToleranceInner / InnerTol
                if (InnerMax_v == null)
                {
                    var innerTol = ExtractTolerance(spsNew.ToleranceInner) ?? ExtractTolerance(spsNew.InnerTol);
                    // Try ToleranceInner itself first (may contain full "8.0±0.2")
                    var innerSrc = spsNew.ToleranceInner?.Contains("±") == true
                        ? spsNew.ToleranceInner
                        : spsNew.InnerTarget;
                    (InnerMin_v, InnerMax_v, InnerTarget_v) = ParseRange(innerSrc, innerTol);
                }

                // Level 3 (final fallback): use Dimensi field — same as legacy Masterlist
                if (InnerMax_v == null)
                {
                    var innerTolFallback = ExtractTolerance(spsNew.ToleranceInner) ?? ExtractTolerance(spsNew.InnerTol);
                    (InnerMin_v, InnerMax_v, InnerTarget_v) = ParseRange(spsNew.Dimensi, innerTolFallback);
                    if (InnerMax_v != null)
                        Console.WriteLine($"[GetChartData][NEW] Inner resolved from Dimensi field: {spsNew.Dimensi}");
                }

                var InnerLCL_v = ParseStandardValue(spsNew.InnerLCL);
                var InnerUCL_v = ParseStandardValue(spsNew.InnerUCL);

                decimal? ThickMin_v, ThickMax_v, ThickTarget_v;
                (ThickMin_v, ThickMax_v, ThickTarget_v) = ParseRange(spsNew.ThickMin ?? spsNew.ThickTarget);
                if (ThickMin_v == null)
                {
                    var tol = ExtractTolerance(spsNew.ThickTol);
                    (ThickMin_v, ThickMax_v, ThickTarget_v) = ParseRange(spsNew.ThickTarget ?? spsNew.TebalInner, tol);
                }

                decimal? TotalMin_v, TotalMax_v, TotalTarget_v;
                (TotalMin_v, TotalMax_v, TotalTarget_v) = ParseRange(spsNew.TotalMin ?? spsNew.TotalTarget);
                if (TotalMin_v == null)
                {
                    var tol = ExtractTolerance(spsNew.TotalTol);
                    (TotalMin_v, TotalMax_v, TotalTarget_v) = ParseRange(spsNew.TotalTarget ?? spsNew.TebalTotal, tol);
                }

                Console.WriteLine($"[GetChartData][NEW] Inner: target={InnerTarget_v}, min={InnerMin_v}, max={InnerMax_v}");
                Console.WriteLine($"[GetChartData][NEW] Thick: target={ThickTarget_v}, min={ThickMin_v}, max={ThickMax_v}");
                Console.WriteLine($"[GetChartData][NEW] Total: target={TotalTarget_v}, min={TotalMin_v}, max={TotalMax_v}");

                dimensionStandard = new
                {
                    innerDiameter = new {
                        target = InnerTarget_v,
                        min = InnerMin_v,
                        max = InnerMax_v,
                        lcl = InnerLCL_v,
                        ucl = InnerUCL_v
                    },
                    innerThickness = new {
                        target = ThickTarget_v,
                        min = ThickMin_v,
                        max = ThickMax_v
                    },
                    totalThickness = new {
                        target = TotalTarget_v,
                        min = TotalMin_v,
                        max = TotalMax_v
                    }
                };

                spsPayload = new {
                    HeadTemp1 = new { target = spsNew.HeadTemp1_Asli, min = spsNew.HeadTemp1_Min, max = spsNew.HeadTemp1_Max },
                    HeadTemp2 = new { target = spsNew.HeadTemp2_Asli, min = spsNew.HeadTemp2_Min, max = spsNew.HeadTemp2_Max },
                    HeadTemp3 = new { target = spsNew.HeadTemp3_Asli, min = spsNew.HeadTemp3_Min, max = spsNew.HeadTemp3_Max },
                    Cylinder1_1 = new { target = spsNew.Cylinder1_1_Asli, min = spsNew.Cylinder1_1_Min, max = spsNew.Cylinder1_1_Max },
                    Cylinder1_2 = new { target = spsNew.Cylinder1_2_Asli, min = spsNew.Cylinder1_2_Min, max = spsNew.Cylinder1_2_Max },
                    Cylinder1_3 = new { target = spsNew.Cylinder1_3_Asli, min = spsNew.Cylinder1_3_Min, max = spsNew.Cylinder1_3_Max },
                    Cylinder2_1 = new { target = spsNew.Cylinder2_1_Asli, min = spsNew.Cylinder2_1_Min, max = spsNew.Cylinder2_1_Max },
                    Cylinder2_2 = new { target = spsNew.Cylinder2_2_Asli, min = spsNew.Cylinder2_2_Min, max = spsNew.Cylinder2_2_Max },
                    Cylinder2_3 = new { target = spsNew.Cylinder2_3_Asli, min = spsNew.Cylinder2_3_Min, max = spsNew.Cylinder2_3_Max },
                    Cylinder3_1 = new { target = spsNew.Cylinder3_1_Asli, min = spsNew.Cylinder3_1_Min, max = spsNew.Cylinder3_1_Max },
                    Cylinder3_2 = new { target = spsNew.Cylinder3_2_Asli, min = spsNew.Cylinder3_2_Min, max = spsNew.Cylinder3_2_Max },
                    Cylinder3_3 = new { target = spsNew.Cylinder3_3_Asli, min = spsNew.Cylinder3_3_Min, max = spsNew.Cylinder3_3_Max },
                    ScrewTemp1 = new { target = spsNew.ScrewTemp1_Asli, min = spsNew.ScrewTemp1_Min, max = spsNew.ScrewTemp1_Max },
                    ScrewTemp2 = new { target = spsNew.ScrewTemp2_Asli, min = spsNew.ScrewTemp2_Min, max = spsNew.ScrewTemp2_Max },
                    ScrewTemp3 = new { target = spsNew.ScrewTemp3_Asli, min = spsNew.ScrewTemp3_Min, max = spsNew.ScrewTemp3_Max },
                    ScrewSpeed1 = new { target = spsNew.ScrewSpeed1_Asli, min = spsNew.ScrewSpeed1_Min, max = spsNew.ScrewSpeed1_Max },
                    ScrewSpeed2 = new { target = spsNew.ScrewSpeed2_Asli, min = spsNew.ScrewSpeed2_Min, max = spsNew.ScrewSpeed2_Max },
                    ScrewSpeed3 = new { target = spsNew.ScrewSpeed3_Asli, min = spsNew.ScrewSpeed3_Min, max = spsNew.ScrewSpeed3_Max },
                    Pressure1 = new { target = spsNew.Pressure1_Asli, min = spsNew.Pressure1_Min, max = spsNew.Pressure1_Max },
                    Pressure2 = new { target = spsNew.Pressure2_Asli, min = spsNew.Pressure2_Min, max = spsNew.Pressure2_Max },
                    Pressure3 = new { target = spsNew.Pressure3_Asli, min = spsNew.Pressure3_Min, max = spsNew.Pressure3_Max },
                    FeedRollRatio1 = new { target = spsNew.FeedRollRatio1_Asli, min = spsNew.FeedRollRatio1_Min, max = spsNew.FeedRollRatio1_Max },
                    FeedRollRatio2 = new { target = spsNew.FeedRollRatio2_Asli, min = spsNew.FeedRollRatio2_Min, max = spsNew.FeedRollRatio2_Max },
                    FeedRollRatio3 = new { target = spsNew.FeedRollRatio3_Asli, min = spsNew.FeedRollRatio3_Min, max = spsNew.FeedRollRatio3_Max },
                    Feed1 = new { target = spsNew.Feed1_Asli, min = spsNew.Feed1_Min, max = spsNew.Feed1_Max },
                    Feed2 = new { target = spsNew.Feed2_Asli, min = spsNew.Feed2_Min, max = spsNew.Feed2_Max },
                    Feed3 = new { target = spsNew.Feed3_Asli, min = spsNew.Feed3_Min, max = spsNew.Feed3_Max },
                    HoseSpeed = new { target = spsNew.HoseSpeed_Asli, min = spsNew.HoseSpeed_Min, max = spsNew.HoseSpeed_Max },
                    SpiralSpeed = new { target = spsNew.SpiralSpeed_Asli, min = spsNew.SpiralSpeed_Min, max = spsNew.SpiralSpeed_Max },
                    SpiralPitchSetting = new { target = spsNew.SpiralPitchSetting_Asli, min = spsNew.SpiralPitchSetting_Min, max = spsNew.SpiralPitchSetting_Max },
                    SpiralPitchDisplay = new { target = spsNew.SpiralPitchDisplay_Asli, min = spsNew.SpiralPitchDisplay_Min, max = spsNew.SpiralPitchDisplay_Max },
                    PresetValue = new { target = spsNew.PresetValue_Asli, min = spsNew.PresetValue_Min, max = spsNew.PresetValue_Max },
                    ControlValue = new { target = spsNew.ControlValue_Asli, min = spsNew.ControlValue_Min, max = spsNew.ControlValue_Max },
                    ChillerWaterTemp = new { target = spsNew.ChillerWaterTemp_Asli, min = spsNew.ChillerWaterTemp_Min, max = spsNew.ChillerWaterTemp_Max },
                    CaterpillarGap = new { target = spsNew.CaterpillarGap_Asli, min = spsNew.CaterpillarGap_Min, max = spsNew.CaterpillarGap_Max },
                    TakeupConveyorSpeed = new { target = spsNew.TakeUpConveyorSpeed_Asli, min = spsNew.TakeUpConveyorSpeed_Min, max = spsNew.TakeUpConveyorSpeed_Max },
                    CoolConveyorSpeed = new { target = spsNew.CoolConveyorSpeed_Asli, min = spsNew.CoolConveyorSpeed_Min, max = spsNew.CoolConveyorSpeed_Max },
                    ConveyorRatio = new { target = spsNew.ConveyorRatio_Asli, min = spsNew.ConveyorRatio_Min, max = spsNew.ConveyorRatio_Max },
                    UnsmoothSurface = spsNew.UnsmoothSurface,
                    ItemList = spsNew.DocumentNumber
                };
            }
            else if (spsOld != null)
            {
                Console.WriteLine($"[GetChartData] Using MasterlistSpsDoubleLayers (LEGACY) for standards");
                
                // Standardize with Legacy Logic
                var innerTolValue = ExtractTolerance(spsOld.ToleranceInner) ?? ExtractTolerance(spsOld.InnerTol);
                var innerDimSource = spsOld.ToleranceInner?.Contains("±") == true ? spsOld.ToleranceInner : spsOld.Dimensi;
                var (innerMin, innerMax, innerTarget) = ParseRange(innerDimSource, innerTolValue);

                var innerLcl = ParseStandardValue(spsOld.InnerLCL);
                var innerUcl = ParseStandardValue(spsOld.InnerUCL);

                var thickTolValue = ExtractTolerance(spsOld.ThickTol);
                var (innerThickMin, innerThickMax, innerThickTarget) = ParseRange(spsOld.TebalInner, thickTolValue);
                
                var totalTolValue = ExtractTolerance(spsOld.TotalTol);
                var (totalMin, totalMax, totalTarget) = ParseRange(spsOld.TebalTotal, totalTolValue);

                dimensionStandard = new
                {
                    innerDiameter = new { target = innerTarget, min = innerMin, max = innerMax, lcl = innerLcl, ucl = innerUcl },
                    innerThickness = new { target = innerThickTarget, min = innerThickMin, max = innerThickMax },
                    totalThickness = new { target = totalTarget, min = totalMin, max = totalMax }
                };

                var s = spsOld;
                spsPayload = new {
                    HeadTemp1 = ParseRange(s.HeadTemp1, 5),
                    HeadTemp2 = ParseRange(s.HeadTemp2, 5),
                    HeadTemp3 = ParseRange(s.HeadTemp3, 5),
                    Cylinder1_1 = ParseRange(s.Cylinder1_1, 5),
                    Cylinder1_2 = ParseRange(s.Cylinder1_2, 5),
                    Cylinder1_3 = ParseRange(s.Cylinder1_3, 5),
                    Cylinder2_1 = ParseRange(s.Cylinder2_1, 5),
                    Cylinder2_2 = ParseRange(s.Cylinder2_2, 5),
                    Cylinder2_3 = ParseRange(s.Cylinder2_3, 5),
                    Cylinder3_1 = ParseRange(s.Cylinder3_1, 5),
                    Cylinder3_2 = ParseRange(s.Cylinder3_2, 5),
                    Cylinder3_3 = ParseRange(s.Cylinder3_3, 5),
                    ScrewTemp1 = ParseRange(s.ScrewTemp1, 5),
                    ScrewTemp2 = ParseRange(s.ScrewTemp2, 5),
                    ScrewTemp3 = ParseRange(s.ScrewTemp3, 5),
                    ScrewSpeed1 = ParseRange(s.ScrewSpeed1, 2),
                    ScrewSpeed2 = ParseRange(s.ScrewSpeed2, 2),
                    ScrewSpeed3 = ParseRange(s.ScrewSpeed3, 2),
                    Pressure1 = ParseRange(s.Pressure1, 1),
                    Pressure2 = ParseRange(s.Pressure2, 1),
                    Pressure3 = ParseRange(s.Pressure3, 1),
                    FeedRollRatio1 = ParseRange(s.FeedRollRatio1),
                    FeedRollRatio2 = ParseRange(s.FeedRollRatio2),
                    FeedRollRatio3 = ParseRange(s.FeedRollRatio3),
                    Feed1 = ParseRange(s.Feed1),
                    Feed2 = ParseRange(s.Feed2),
                    Feed3 = ParseRange(s.Feed3),
                    HoseSpeed = ParseRange(s.HoseSpeed),
                    SpiralSpeed = ParseRange(s.SpiralSpeed),
                    SpiralPitchSetting = ParseRange(s.SpiralPitchSetting),
                    SpiralPitchDisplay = ParseRange(s.SpiralPitchDisplay),
                    PresetValue = ParseRange(s.PresetValue),
                    ControlValue = ParseRange(s.ControlValue),
                    ChillerWaterTemp = ParseRange(s.ChillerWaterTemp, 3),
                    CaterpillarGap = ParseRange(s.CaterpillarGap),
                    TakeupConveyorSpeed = ParseRange(s.TakeUpConveyorSpeed),
                    CoolConveyorSpeed = ParseRange(s.CoolConveyorSpeed),
                    ConveyorRatio = ParseRange(s.ConveyorRatio),
                    UnsmoothSurface = ParseRange(s.UnsmoothSurface),
                    ItemList = s.DocumentNumber
                };
            }

            Console.WriteLine($"[GetChartData] === FINAL JSON RESPONSE ===");
            
            // Smart Fallback for VIN/Item Code display
            string displayVin = report.VinCode ?? "";
            string displayItem = report.ItemCode ?? "";
            
            // Only fallback to ItemCode if the report has no item/vin AND we have a direct SPS match
            if (string.IsNullOrEmpty(displayItem) && !string.IsNullOrEmpty(effectiveItemCode))
            {
                displayItem = effectiveItemCode;
            }

            if (string.IsNullOrEmpty(displayVin) && string.IsNullOrEmpty(displayItem) && hasSps)
            {
                if (spsNew != null) displayItem = spsNew.DocumentNumber ?? "";
                else if (spsOld != null) displayItem = spsOld.DocumentNumber ?? "";
            }

            Console.WriteLine($"[GetChartData] VinCode: {displayVin}");
            Console.WriteLine($"[GetChartData] ItemCode: {displayItem}");
            Console.WriteLine($"[GetChartData] SPS Found: {hasSps}");
            
            return Json(new {
                success = true,
                report = new {
                    id = report.Id,
                    documentNumber = report.DocumentNumber,
                    hoseType = report.HoseType,
                    vinCode = displayVin,
                    itemCode = displayItem,
                    machineName = report.MachineName,
                    dimension = report.Dimension,
                    customerName = report.CustomerName,
                    shift = report.Shift,
                    productionDate = report.ProductionDate.ToString("yyyy-MM-dd"),
                    status = report.Status
                },
                spsFound = hasSps,
                spsStandard = spsPayload,
                initials = new {
                    report.InitHeadTempInner, report.InitHeadTempOuter,
                    report.InitCylinder1TempInner, report.InitCylinder1TempOuter,
                    report.InitCylinder2TempInner, report.InitCylinder2TempOuter,
                    report.InitCylinder3TempInner, report.InitCylinder3TempOuter,
                    report.InitScrewTempInner, report.InitScrewTempOuter,
                    report.InitScrewSpeedInner, report.InitScrewSpeedOuter,
                    report.InitFeedRollRatioInner, report.InitFeedRollRatioOuter,
                    report.InitPressureInner, report.InitPressureOuter,
                    report.InitHoseSpeed, report.InitSpiralSpeed,
                    report.InitSpiralPitchSetting, report.InitSpiralPitchDisplay,
                    report.InitPresetValue, report.InitControlValue,
                    report.InitChillerWaterTemp, report.InitCaterpillarGap,
                    report.InitTakeupConveyorSpeed, report.InitCoolConveyorSpeed,
                    report.InitUnsmoothSurface
                },
                readings,
                dimensionSeries,
                dimensionStandard,
                sensorData
            });
        }

        private static bool ContainsAny(string? source, params string[] terms)
        {
            if (string.IsNullOrWhiteSpace(source)) return false;
            var value = source.Trim().ToUpperInvariant();
            return terms.Any(t => value.Contains(t.ToUpperInvariant()));
        }

        private static IEnumerable<(string label, decimal value)> ExpandDimensionReadings(DimensionMeasurement measurement, DateTime productionDate)
        {
            // Handle multiple times if they are stored as a comma-separated string (e.g., from Input Dimensi Timeline)
            var times = (measurement.TimeSection ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries);
            var dateLabel = productionDate.ToString("dd MMM");
            
            var values = new[]
            {
                measurement.R1,
                measurement.R2,
                measurement.R3,
                measurement.R4,
                measurement.R5
            };

            for (var index = 0; index < values.Length; index++)
            {
                if (values[index].HasValue)
                {
                    // Pick the specific time for this reading index if available, 
                    // otherwise fallback to the last time or recorded time.
                    string exactTime;
                    if (index < times.Length)
                    {
                        exactTime = times[index].Trim();
                    }
                    else
                    {
                        exactTime = times.Length > 0 
                            ? times.Last().Trim() 
                            : (string.IsNullOrWhiteSpace(measurement.TimeSection) ? measurement.RecordedTime.ToString("HH:mm") : measurement.TimeSection);
                    }

                    yield return ($"{dateLabel} {exactTime} • R{index + 1}", values[index]!.Value);
                }
            }
        }

        private static decimal? ParseStandardValue(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            var normalized = raw.Replace(',', '.');
            var match = System.Text.RegularExpressions.Regex.Match(normalized, @"-?\d+(?:\.\d+)?");
            if (!match.Success) return null;
            return decimal.TryParse(match.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var value)
                ? value
                : null;
        }

        private static (decimal? min, decimal? max, decimal? target) ParseRange(object? rawObj, decimal? defaultTolerance = null)
        {
            if (rawObj == null) return (null, null, null);
            string? raw = rawObj.ToString();
            if (string.IsNullOrWhiteSpace(raw)) return (null, null, null);

            var normalized = raw.Replace(',', '.');

            // 1. Explicit ± format (e.g. "1.8±0.2")
            if (normalized.Contains('±'))
            {
                var mPlusMinus = System.Text.RegularExpressions.Regex.Match(normalized, @"(-?\d+(?:\.\d+)?)\s*±\s*(-?\d+(?:\.\d+)?)");
                if (mPlusMinus.Success)
                {
                    if (decimal.TryParse(mPlusMinus.Groups[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v) &&
                        decimal.TryParse(mPlusMinus.Groups[2].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var tol))
                    {
                        return (v - Math.Abs(tol), v + Math.Abs(tol), v);
                    }
                }
            }

            // 1.5. If string contains 'x' or '×', it's a dimension (e.g. "16 x 24"), not a range.
            // First number is Inner Diameter.
            if (normalized.Contains("x", StringComparison.OrdinalIgnoreCase) || normalized.Contains('×'))
            {
                var matchesDim = System.Text.RegularExpressions.Regex.Matches(normalized, @"-?\d+(?:\.\d+)?");
                if (matchesDim.Count > 0 && decimal.TryParse(matchesDim[0].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                {
                    if (defaultTolerance.HasValue)
                        return (v - Math.Abs(defaultTolerance.Value), v + Math.Abs(defaultTolerance.Value), v);
                    return (v, v, v);
                }
            }

            var matches = System.Text.RegularExpressions.Regex.Matches(normalized, @"-?\d+(?:\.\d+)?");

            // 2. Range format (e.g. "1.6 - 2.0")
            if (matches.Count >= 2)
            {
                if (decimal.TryParse(matches[0].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v1) &&
                    decimal.TryParse(matches[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v2))
                {
                    return (Math.Min(v1, v2), Math.Max(v1, v2), (v1 + v2) / 2);
                }
            }

            // 3. Single Target format (e.g. "1.8")
            if (matches.Count == 1)
            {
                if (decimal.TryParse(matches[0].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var v))
                {
                    if (defaultTolerance.HasValue)
                        return (v - Math.Abs(defaultTolerance.Value), v + Math.Abs(defaultTolerance.Value), v);
                    return (v, v, v);
                }
            }

            return (null, null, null);
        }

        private static decimal? ExtractTolerance(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            var normalized = raw.Replace(',', '.');
            
            // Check if it explicitly contains ±
            if (normalized.Contains('±'))
            {
                var mPlusMinus = System.Text.RegularExpressions.Regex.Match(normalized, @"±\s*(-?\d+(?:\.\d+)?)");
                if (mPlusMinus.Success && decimal.TryParse(mPlusMinus.Groups[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var tol))
                {
                    return Math.Abs(tol);
                }
            }
            
            // Otherwise try to grab the last number, assuming it's a pure tolerance field like "0.2" or "+-0.2"
            var matches = System.Text.RegularExpressions.Regex.Matches(normalized, @"-?\d+(?:\.\d+)?");
            if (matches.Count > 0 && decimal.TryParse(matches[^1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var lastNum))
            {
                var absVal = Math.Abs(lastNum);
                // Sanity check: tolerances are usually small. If > 5.0, it might be a target value.
                if (absVal <= 5.0m) return absVal;
            }
            return null;
        }

        private static bool IsPlaceholderItemCode(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return true;
            var normalized = value.Trim().Replace(" ", "");
            // Filter placeholder values AND pure numeric codes (incomplete planning data)
            return normalized is "-" or "--" or "---" or "N/A" or "NA" 
                || normalized.All(char.IsDigit);  // Filter: 3110, 1290, etc.
        }

        // GET: ProductionReport/GetDocumentList — for autocomplete search
        [HttpGet]
        public async Task<IActionResult> GetDocumentList(string? q = null)
        {
            var query = _context.ProductionReports.AsQueryable();

            if (!string.IsNullOrWhiteSpace(q))
            {
                var search = q.Trim().ToUpper();
                query = query.Where(r =>
                    (r.DocumentNumber != null && r.DocumentNumber.ToUpper().Contains(search)) ||
                    (r.VinCode != null && r.VinCode.ToUpper().Contains(search)) ||
                    (r.HoseType != null && r.HoseType.ToUpper().Contains(search)));
            }

            var docs = await query
                .OrderByDescending(r => r.ProductionReadings.Count)
                .ThenByDescending(r => r.CreatedDate)
                .Take(30)
                .Select(r => new {
                    r.DocumentNumber,
                    r.HoseType,
                    r.VinCode,
                    r.MachineName,
                    date = r.ProductionDate.ToString("dd MMM yyyy"),
                    readingsCount = r.ProductionReadings.Count
                })
                .ToListAsync();

            return Json(docs);
        }

        // GET: ProductionReport/GetDocumentsByItem — get production reports by SPS ItemCode
        // Fallback chain: ItemCode/VinCode → HoseType from SPS → All documents
        [HttpGet]
        public async Task<IActionResult> GetDocumentsByItem(string? itemCode = null, string? startDate = null, string? endDate = null, string? shift = null, string? hoseType = null, string? pic = null)
        {
            // Only show production reports that have corresponding DimensionReports
            var dimDocNumbers = await _context.DimensionReports.Select(dr => dr.DocumentNumber).Distinct().ToListAsync();
            var dimItemCodes = await _context.DimensionReports.Select(dr => dr.ItemCode).Distinct().ToListAsync();

            IQueryable<ProductionReport> query = _context.ProductionReports
                .Where(r => dimDocNumbers.Contains(r.DocumentNumber) || dimItemCodes.Contains(r.ItemCode));

            // Apply Filters (Date Range) - Using direct comparison for better EF compatibility
            if (!string.IsNullOrWhiteSpace(startDate))
            {
                if (DateTime.TryParse(startDate, out DateTime sDate))
                {
                    var startOfDay = sDate.Date;
                    query = query.Where(r => r.ProductionDate >= startOfDay);
                }
            }
            if (!string.IsNullOrWhiteSpace(endDate))
            {
                if (DateTime.TryParse(endDate, out DateTime eDate))
                {
                    var endOfDay = eDate.Date.AddDays(1).AddTicks(-1);
                    query = query.Where(r => r.ProductionDate <= endOfDay);
                }
            }

            if (!string.IsNullOrWhiteSpace(shift) && !shift.Equals("SHIFT ALL", StringComparison.OrdinalIgnoreCase))
            {
                var shiftNorm = shift.Trim().ToUpper();
                query = query.Where(r => r.Shift != null && r.Shift.Trim().ToUpper().Contains(shiftNorm));
            }

            if (!string.IsNullOrWhiteSpace(hoseType))
            {
                var hup = hoseType.Trim().ToUpper();
                query = query.Where(r => r.HoseType != null && r.HoseType.ToUpper().Contains(hup));
            }

            if (!string.IsNullOrWhiteSpace(pic))
            {
                var pup = pic.Trim().ToUpper();
                query = query.Where(r => r.CreatedBy != null && r.CreatedBy.ToUpper().Contains(pup));
            }

            if (!string.IsNullOrWhiteSpace(itemCode))
            {
                var search = itemCode.Trim().ToUpper();

                // Exact match by ItemCode or VinCode
                var byCode = await query
                    .Where(r =>
                        (r.ItemCode != null && r.ItemCode.ToUpper().Contains(search)) ||
                        (r.VinCode  != null && r.VinCode.ToUpper().Contains(search)))
                    .OrderByDescending(r => r.ProductionReadings.Count)
                    .ThenByDescending(r => r.CreatedDate)
                    .Take(50)
                    .Select(r => new {
                        r.DocumentNumber, r.HoseType, r.VinCode, r.ItemCode,
                        r.MachineName, r.Shift,
                        date = r.ProductionDate.ToString("dd MMM yyyy"),
                        readingsCount = r.ProductionReadings.Count
                    })
                    .ToListAsync();

                // If a specific itemCode is requested, we strictly only return matches for that code.
                // We no longer fallback to matching by HoseType or returning random documents,
                // as that causes confusion (e.g., showing TA1440 data when searching for NA1450).
                return Json(byCode);
            }

            // No itemCode filter provided — return latest 50 docs
            var allDocs = await query
                .OrderByDescending(r => r.ProductionReadings.Count)
                .ThenByDescending(r => r.CreatedDate)
                .Take(50)
                .Select(r => new {
                    r.DocumentNumber, r.HoseType, r.VinCode, r.ItemCode,
                    r.MachineName, r.Shift,
                    date = r.ProductionDate.ToString("dd MMM yyyy"),
                    readingsCount = r.ProductionReadings.Count
                })
                .ToListAsync();

            return Json(allDocs);
        }

        // GET: ProductionReport/GetFilterOptions — Get unique products and PICs for filters
        [HttpGet]
        public async Task<IActionResult> GetFilterOptions(
            string? startDate = null,
            string? endDate = null,
            string? shift = null,
            string? hoseType = null,
            string? itemCode = null)
        {
            IQueryable<ProductionReport> query = _context.ProductionReports;

            if (!string.IsNullOrWhiteSpace(startDate) && DateTime.TryParse(startDate, out var sDate))
                query = query.Where(r => r.ProductionDate >= sDate.Date);

            if (!string.IsNullOrWhiteSpace(endDate) && DateTime.TryParse(endDate, out var eDate))
                query = query.Where(r => r.ProductionDate <= eDate.Date.AddDays(1).AddTicks(-1));

            if (!string.IsNullOrWhiteSpace(shift) && !shift.Equals("SHIFT ALL", StringComparison.OrdinalIgnoreCase))
            {
                var shiftNorm = shift.Trim().ToUpper();
                query = query.Where(r => r.Shift != null && r.Shift.Trim().ToUpper().Contains(shiftNorm));
            }

            if (!string.IsNullOrWhiteSpace(hoseType))
            {
                var hup = hoseType.Trim().ToUpper();
                query = query.Where(r => r.HoseType != null && r.HoseType.ToUpper().Contains(hup));
            }

            if (!string.IsNullOrWhiteSpace(itemCode))
            {
                var iup = itemCode.Trim().ToUpper();
                query = query.Where(r =>
                    (r.ItemCode != null && r.ItemCode.ToUpper().Contains(iup)) ||
                    (r.VinCode != null && r.VinCode.ToUpper().Contains(iup)));
            }

            var hoseTypes = await query
                .Where(r => r.HoseType != null && r.HoseType != "")
                .Select(r => r.HoseType!)
                .Distinct()
                .OrderBy(h => h)
                .ToListAsync();

            var pics = await query
                .Where(r => r.CreatedBy != null && r.CreatedBy != "")
                .Select(r => r.CreatedBy!)
                .Distinct()
                .OrderBy(p => p)
                .ToListAsync();

            return Json(new { hoseTypes, pics });
        }

        [HttpPost]
        public async Task<IActionResult> UpdateProductionTime(int id, DateTime? start, DateTime? end)
        {
            var report = await _context.ProductionReports.FindAsync(id);
            if (report == null) return NotFound(new { success = false, message = "Report not found" });

            report.ProductionStartTime = start;
            report.ProductionEndTime = end;
            await _context.SaveChangesAsync();

            await _hubContext.Clients.All.SendAsync("ReceiveUpdate");
            return Json(new { success = true, message = "Jam produksi berhasil diperbarui." });
        }

        [HttpGet]
        public async Task<IActionResult> GetSensorDataForAnalysis(int reportId, DateTime? start, DateTime? end)
        {
            var report = await _context.ProductionReports.FindAsync(reportId);
            if (report == null) return NotFound("Report not found");

            // Ambil Standar OD
            string targetDocNum = report.DocumentNumber;

            if (!string.IsNullOrEmpty(report.ItemCode))
            {
                var spsList = await _context.SpsItemLists
                    .OrderByDescending(l => l.Id)
                    .FirstOrDefaultAsync(l => l.ItemList != null && l.ItemList.Contains(report.ItemCode));
                
                if (spsList != null && !string.IsNullOrEmpty(spsList.DocumentNumber) && spsList.DocumentNumber != "-")
                {
                    targetDocNum = spsList.DocumentNumber;
                }

                // Fetch SPS record with actual OD values (prefer non-null)
                var spsRecord = await _context.SpsNoDocs
                    .Where(s => s.DocumentNumber == targetDocNum)
                    .FirstOrDefaultAsync(s => s.OdSensor_Asli != null && s.OdSensor_Asli != 0);
                if (spsRecord != null)
                {
                    targetDocNum = spsRecord.DocumentNumber;
                }
            }

            var sps = await _context.SpsNoDocs
                .OrderByDescending(s => s.DocumentNumber)
                .FirstOrDefaultAsync(s => s.DocumentNumber == targetDocNum || s.DocumentNumber == report.DocumentNumber);

            decimal stdOD = sps?.OdSensor_Asli ?? 0;
            decimal upperLimit = sps?.OdSensor_Max ?? 0;
            decimal lowerLimit = sps?.OdSensor_Min ?? 0;

            string sensorMachineCode = report.MachineName ?? "";
            if (sensorMachineCode.Contains("CHS2L")) sensorMachineCode = "CHS 2 Layer";
            else if (sensorMachineCode.Contains("CHS3L")) sensorMachineCode = "CHS 3 Layer";

            var logsQuery = _context.SensorIngestLogs
                .Where(l => (l.MachineCode == report.MachineName || l.MachineCode == sensorMachineCode)
                            && l.MetricType == "outer_diameter");
            
            var (defaultStart, defaultEnd) = ResolveProductionWindow(report);
            var rangeStart = start ?? defaultStart;
            var rangeEnd = end ?? defaultEnd;

            logsQuery = logsQuery.Where(l => l.SensorTimestamp >= rangeStart);
            logsQuery = logsQuery.Where(l => l.SensorTimestamp <= rangeEnd);

            var logs = await logsQuery
                .OrderBy(l => l.SensorTimestamp)
                .Select(l => new {
                    time = l.SensorTimestamp.ToString("HH:mm:ss"),
                    value = l.MetricValue,
                    device = l.DeviceId
                })
                .ToListAsync();

            return Json(new { 
                success = true, 
                stdOD = stdOD, 
                upperLimit = upperLimit, 
                lowerLimit = lowerLimit, 
                logs = logs 
            });
        }

        /// <summary>
        /// Jendela produksi 4 jam per mesin.
        /// Save pertama: Mulai = sekarang, Selesai = sekarang + 4 jam.
        /// Save berikutnya dalam jendela aktif: pakai Mulai/Selesai yang sama.
        /// Setelah jendela habis: save baru buka jendela baru.
        /// </summary>
        private async Task<(DateTime start, DateTime end)> GetProductionWindowAsync(string? machineName)
        {
            var now = DateTime.Now;

            if (!string.IsNullOrWhiteSpace(machineName))
            {
                var activeWindow = await _context.ProductionReports
                    .Where(r => r.MachineName == machineName
                        && r.ProductionStartTime != null
                        && r.ProductionEndTime != null
                        && r.ProductionEndTime > now)
                    .OrderByDescending(r => r.ProductionStartTime)
                    .Select(r => new { r.ProductionStartTime, r.ProductionEndTime })
                    .FirstOrDefaultAsync();

                if (activeWindow?.ProductionStartTime != null && activeWindow.ProductionEndTime != null)
                    return (activeWindow.ProductionStartTime.Value, activeWindow.ProductionEndTime.Value);
            }

            return (now, now.AddHours(4));
        }

        private static (DateTime start, DateTime end) ResolveProductionWindow(ProductionReport report)
        {
            var start = report.ProductionStartTime ?? report.CreatedDate;
            var end = report.ProductionEndTime ?? start.AddHours(4);
            return (start, end);
        }

        [HttpGet]
        public async Task<IActionResult> GetAllOfflineMasterData(DateTime? targetDate = null)
        {
            try
            {
                var date = (targetDate ?? DateTime.Today).Date;

                var cultureID = new System.Globalization.CultureInfo("id-ID");
                
                // 1. Planning Master (Live from ELWP_PRD)
                var planningsQuery = from p in _elwpContext.ElwpPlannings
                                     join m in _elwpContext.ElwpMachines on p.MesinId equals m.Id into machineJoin
                                     from m in machineJoin.DefaultIfEmpty()
                                     where p.TanggalPlanning >= date && p.TanggalPlanning < date.AddDays(1)
                                     select new { p, MachineName = m != null ? m.NamaMesin : "UNKNOWN", MachineCode = m != null ? m.KodeMesin : "" };
                
                var elwpRows = await planningsQuery.ToListAsync();
                
                var plannings = elwpRows.Select(x => new {
                    itemCode = x.p.KodeItem,
                    itemName = (x.p.PartName ?? "#N/A") + (string.IsNullOrEmpty(x.p.PnSap) ? "" : " | " + x.p.PnSap),
                    machineName = x.MachineName,
                    machineCode = x.MachineCode,
                    dateShift = $"{(x.p.TanggalPlanning?.ToString("dddd, d MMMM yyyy", cultureID) ?? "").ToUpper()} SHIFT {x.p.Shift}",
                    date = x.p.TanggalPlanning?.ToString("yyyy-MM-dd"),
                    shift = x.p.Shift?.ToString(),
                    isFiltered = false,
                    hasSps = true
                }).OrderBy(p => p.itemName).ToList();

                // 2. SPS Master
                var spsRaw = await _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .AsNoTracking()
                    .ToListAsync();

                var spsList = new List<object>();
                foreach(var s in spsRaw)
                {
                    spsList.Add(MapMasterToSps(s));
                }

                // 3. Todays Summary (All machines)
                var endDate = date.AddDays(1);
                var paramReports = await _context.ProductionReports
                    .Where(r => r.CreatedDate >= date && r.CreatedDate < endDate)
                    .Select(r => new {
                        r.Id, r.HoseType, r.VinCode, r.Shift, r.CreatedBy, r.Status, r.CreatedDate,
                        r.DocumentNumber, r.RevisionNumber,
                        r.MachineName
                    }).ToListAsync();

                var dimReports = await _context.DimensionReports
                    .Where(r => r.CreatedDate >= date && r.CreatedDate < endDate)
                    .Select(r => new {
                        r.Id, r.HoseType, r.Shift, r.CreatedBy, r.Status, r.CreatedDate,
                        r.DocumentNumber, r.RevisionNumber, r.MachineName
                    }).ToListAsync();
                    
                var summary = new {
                    parameterReports = paramReports,
                    dimensionReports = dimReports
                };

                // 4. Users and Machines for Offline Login
                var users = await _elwpContext.ElwpUsers
                    .Where(u => u.IsActive && u.AreaId == 1)
                    .OrderBy(u => u.FullName)
                    .Select(u => new { u.Id, u.FullName, u.Username })
                    .ToListAsync();

                var machines = await _elwpContext.ElwpMachines
                    .Where(m => m.IsActive && m.AreaId == 1 && m.KodeMesin != "DL01" && m.KodeMesin != "DL02")
                    .OrderBy(m => m.KodeMesin)
                    .Select(m => new { m.Id, m.KodeMesin, m.NamaMesin })
                    .ToListAsync();

                return Json(new { success = true, plannings, spsList, summary, users, machines });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

    }
}


