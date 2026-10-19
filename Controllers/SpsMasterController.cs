using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using VelastoProductionSystem.Data;
using VelastoProductionSystem.Models;
using VelastoProductionSystem.Helpers;
using OfficeOpenXml;
using System.Text;
using System.Reflection;
using System.ComponentModel.DataAnnotations;
using VelastoProductionSystem.Services;
using System.Text.Json;
using System.Security.Cryptography;

namespace VelastoProductionSystem.Controllers
{
    public class SpsMasterController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ILogger<SpsMasterController> _logger;
        private readonly IApprovalService _approvalService;

        public SpsMasterController(ApplicationDbContext context, ILogger<SpsMasterController> logger, IApprovalService approvalService)
        {
            _context = context;
            _logger = logger;
            _approvalService = approvalService;
            
            // EPPlus license sudah di-set secara global di Program.cs
        }

        // ==================== HELPER METHOD: PARSE Â± FORMAT ====================
        /// <summary>
        /// Parse nilai dengan format "XÂ±Y", "Max X", "Min X", atau single value
        /// Contoh: "4.3Â±0.1" â†’ Min=4.2, Asli=4.3, Max=4.4
        ///         "Max 120" â†’ Min=null, Asli=null, Max=120
        ///         "73" â†’ Min=null, Asli=73, Max=null
           private (decimal? min, decimal? asli, decimal? max) ParsePlusMinusValue(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                _logger.LogInformation($"ParsePlusMinusValue: Input is NULL or empty");
                return (null, null, null);
            }

            try
            {
                // Clean input: normalize decimal separator (comma to period) and remove spaces
                var cleaned = input.Trim().Replace(" ", "").Replace(",", ".");
                _logger.LogInformation($"ParsePlusMinusValue: Input=[{input}], Cleaned=[{cleaned}]");

                // Check for asymmetric tolerance format: e.g. "10.5+0.1/-0.2", "10.5(+0.1/-0.2)", "10.5+0.1-0.2"
                var asymMatch = System.Text.RegularExpressions.Regex.Match(cleaned,
                    @"^([0-9]+(?:\.[0-9]+)?)\s*(?:\()?[\+]+\s*([0-9]+(?:\.[0-9]+)?)\s*[\/\-]+\s*[\-]?\s*([0-9]+(?:\.[0-9]+)?)(?:\))?$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                if (asymMatch.Success)
                {
                    if (decimal.TryParse(asymMatch.Groups[1].Value, System.Globalization.NumberStyles.Any, 
                            System.Globalization.CultureInfo.InvariantCulture, out decimal baseVal) &&
                        decimal.TryParse(asymMatch.Groups[2].Value, System.Globalization.NumberStyles.Any, 
                            System.Globalization.CultureInfo.InvariantCulture, out decimal upperTol) &&
                        decimal.TryParse(asymMatch.Groups[3].Value, System.Globalization.NumberStyles.Any, 
                            System.Globalization.CultureInfo.InvariantCulture, out decimal lowerTol))
                    {
                        var min = baseVal - lowerTol;
                        var max = baseVal + upperTol;
                        _logger.LogInformation($"✅ Parsed asymmetric '{input}' → Min={min}, Asli={baseVal}, Max={max}");
                        return (min, baseVal, max);
                    }
                }
                
                // Check for "Max X" format (case insensitive)
                if (cleaned.StartsWith("max", StringComparison.OrdinalIgnoreCase) || 
                    cleaned.StartsWith("maks", StringComparison.OrdinalIgnoreCase))
                {
                    // Skip "max", "maks", and any dots/spaces to find the first digit
                    var numStart = 0;
                    for (int i = 0; i < cleaned.Length; i++)
                    {
                        if (char.IsDigit(cleaned[i]))
                        {
                            numStart = i;
                            break;
                        }
                    }
                    
                    if (numStart > 0)
                    {
                        var numPart = cleaned.Substring(numStart);
                        if (decimal.TryParse(numPart, System.Globalization.NumberStyles.Any, 
                            System.Globalization.CultureInfo.InvariantCulture, out decimal maxValue))
                        {
                            _logger.LogInformation($"Parsed '{input}' as Max format → Max={maxValue}");
                            return (null, null, maxValue);
                        }
                    }
                }
                
                // Check for "Min X" format (case insensitive)
                if (cleaned.StartsWith("min", StringComparison.OrdinalIgnoreCase))
                {
                    // Skip "min" and any dots/spaces to find the first digit
                    var numStart = 0;
                    for (int i = 0; i < cleaned.Length; i++)
                    {
                        if (char.IsDigit(cleaned[i]))
                        {
                            numStart = i;
                            break;
                        }
                    }
                    
                    if (numStart > 0)
                    {
                        var numPart = cleaned.Substring(numStart);
                        if (decimal.TryParse(numPart, System.Globalization.NumberStyles.Any, 
                            System.Globalization.CultureInfo.InvariantCulture, out decimal minValue))
                        {
                            _logger.LogInformation($"Parsed '{input}' as Min format → Min={minValue}");
                            return (minValue, null, null);
                        }
                    }
                }
                
                // Check for ± symbol (or alternative: +/-, +-, ~, etc.)
                // IMPORTANT: Include both the real Unicode ± (U+00B1) and common encoding variants
                var plusMinusSymbols = new[] { "\u00B1", "±", "Â±", "+/-", "+-", "~" };
                
                foreach (var symbol in plusMinusSymbols)
                {
                    if (cleaned.Contains(symbol))
                    {
                        _logger.LogInformation($"ParsePlusMinusValue: Found symbol '{symbol}' (U+{(int)symbol[0]:X4}) in cleaned value");
                        var parts = cleaned.Split(new[] { symbol }, StringSplitOptions.None);
                        _logger.LogInformation($"ParsePlusMinusValue: Split into {parts.Length} parts: [{string.Join("] [", parts)}]");
                        
                        if (parts.Length == 2)
                        {
                            if (decimal.TryParse(parts[0], System.Globalization.NumberStyles.Any, 
                                    System.Globalization.CultureInfo.InvariantCulture, out decimal standard) && 
                                decimal.TryParse(parts[1], System.Globalization.NumberStyles.Any, 
                                    System.Globalization.CultureInfo.InvariantCulture, out decimal tolerance))
                            {
                                var min = standard - tolerance;
                                var max = standard + tolerance;
                                
                                _logger.LogInformation($"✅ Parsed '{input}' → Min={min}, Asli={standard}, Max={max}");
                                return (min, standard, max);
                            }
                            else
                            {
                                _logger.LogWarning($"❌ Failed to parse parts as decimals: [{parts[0]}] and [{parts[1]}]");
                            }
                        }
                    }
                }
                
                // If no special format, try to parse as single value (Asli only)
                if (decimal.TryParse(cleaned, System.Globalization.NumberStyles.Any, 
                    System.Globalization.CultureInfo.InvariantCulture, out decimal singleValue))
                {
                    _logger.LogInformation($"Parsed '{input}' as single value → Asli={singleValue}");
                    return (null, singleValue, null);
                }
                
                _logger.LogWarning($"⚠️ Could not parse '{input}' in any format");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"❌ Exception parsing '{input}': {ex.Message}");
            }
            
            return (null, null, null);
        }

        /// <summary>
        /// Helper untuk assign parsed Â± values ke properties
        /// </summary>
        private void AssignParsedValue(string cellValue, 
            Action<decimal?> setMin, 
            Action<decimal?> setAsli, 
            Action<decimal?> setMax)
        {
            var parsed = ParsePlusMinusValue(cellValue);
            setMin(parsed.min);
            setAsli(parsed.asli);
            setMax(parsed.max);
        }

        public async Task<IActionResult> Index(string? search, string? machine)
        {
            // Ambil semua SpsNoDoc dengan ItemLists - FILTER NO.DOC YANG VALID
            var query = _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .Where(s => 
                    !string.IsNullOrWhiteSpace(s.DocumentNumber) && 
                    s.DocumentNumber != "-" &&
                    !s.DocumentNumber.StartsWith("#REF") &&
                    !s.DocumentNumber.StartsWith("#N/A") &&
                    !s.DocumentNumber.StartsWith("#VALUE") &&
                    !s.DocumentNumber.StartsWith("#ERROR") &&
                    !s.DocumentNumber.StartsWith("ROW-") &&
                    s.DocumentNumber.Length >= 3)
                .AsQueryable();

            // Filter by machine if specified
            if (!string.IsNullOrWhiteSpace(machine))
            {
                var machineUpper = machine.ToUpper();
                query = query.Where(s => 
                    (s.MachineCode != null && s.MachineCode.ToUpper() == machineUpper) ||
                    (s.Machine != null && s.Machine.ToUpper() == machineUpper));
            }

            // Filter by search keyword if specified
            if (!string.IsNullOrWhiteSpace(search))
            {
                var searchUpper = search.ToUpper();
                query = query.Where(s => 
                    (s.DocumentNumber != null && s.DocumentNumber.ToUpper().Contains(searchUpper)) ||
                    (s.HoseType != null && s.HoseType.ToUpper().Contains(searchUpper)) ||
                    (s.Customer != null && s.Customer.ToUpper().Contains(searchUpper)) ||
                    s.ItemLists.Any(i => i.ItemList != null && i.ItemList.ToUpper().Contains(searchUpper)));
            }

            var spsNoDocs = await query
                .OrderByDescending(s => s.IsActive)
                .ThenBy(s => s.DocumentNumber)
                .ToListAsync();

            var expandedData = new List<SpsMaster>();
            foreach (var nodoc in spsNoDocs)
            {
                expandedData.Add(SpsMapper.ToSpsMaster(nodoc, ""));
            }

            // Prepare filter dropdowns and autocomplete lists
            var allData = await _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .Where(s => !string.IsNullOrWhiteSpace(s.DocumentNumber))
                .ToListAsync();

            var machinesList = allData
                .Select(m => m.MachineCode ?? m.Machine)
                .Where(m => !string.IsNullOrEmpty(m))
                .Distinct()
                .OrderBy(m => m)
                .ToList();

            ViewBag.Machines = machinesList;
            ViewBag.AuthorizedMachines = machinesList;  // For filter dropdown in view

            // Prepare ItemList for autocomplete (dari SpsItemLists)
            ViewBag.ItemList = allData
                .SelectMany(s => s.ItemLists ?? new List<SpsItemList>())
                .Select(i => i.ItemList?.Replace("-", ""))
                .Where(i => !string.IsNullOrEmpty(i))
                .Distinct()
                .OrderBy(i => i)
                .ToList();

            // Prepare DocumentNumber list for autocomplete
            ViewBag.DocList = allData
                .Select(s => s.DocumentNumber)
                .Where(d => !string.IsNullOrEmpty(d))
                .Distinct()
                .OrderBy(d => d)
                .ToList();

            ViewBag.Search = search;
            ViewBag.Machine = machine;
            ViewBag.RecordCount = expandedData.Count;

            return View(expandedData);
        }

        public async Task<IActionResult> Details(string? documentNumber)
        {
            if (string.IsNullOrWhiteSpace(documentNumber)) return NotFound();

            var spsNoDoc = await _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .FirstOrDefaultAsync(s => s.DocumentNumber == documentNumber);
            
            if (spsNoDoc == null) return NotFound();

            return View(spsNoDoc);
        }

        [HttpGet]
        public async Task<IActionResult> Create(int? requestId)
        {
            if (requestId.HasValue)
            {
                var currentUser = HttpContext.Session.GetString("UserName") ?? string.Empty;
                var request = await _context.ApprovalRequests
                    .AsNoTracking()
                    .FirstOrDefaultAsync(r => r.Id == requestId.Value
                        && r.ActionType == ApprovalActionType.SpsDocumentCreate
                        && r.RequesterUserName.ToUpper() == currentUser.ToUpper()
                        && !string.IsNullOrWhiteSpace(r.PayloadJson));

                if (request != null)
                {
                    try
                    {
                        var modelFromRequest = JsonSerializer.Deserialize<SpsMaster>(request.PayloadJson!);
                        if (modelFromRequest != null)
                        {
                            ViewBag.SourceRequestId = request.Id;
                            return View(modelFromRequest);
                        }
                    }
                    catch
                    {
                        // Fallback to empty model if payload is not readable.
                    }
                }
            }

            return View(new SpsMaster());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(SpsMaster model, bool saveAsDraft = false, int? sourceRequestId = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(model.DocumentNumber))
                {
                    TempData["ErrorMessage"] = "No. Document tidak boleh kosong";
                    ModelState.AddModelError("DocumentNumber", "No. Document tidak boleh kosong");
                    return View(model);
                }

                if (_approvalService.IsRequesterRole())
                {
                    var targetKey = model.DocumentNumber;

                    if (saveAsDraft)
                    {
                        await _approvalService.SaveDraftRequestAsync(
                            ApprovalActionType.SpsDocumentCreate,
                            targetKey,
                            $"Draft create SPS Document {model.DocumentNumber}",
                            Url.Action(nameof(Create), "SpsMaster") ?? "/SpsMaster/Create",
                            JsonSerializer.Serialize(model),
                            sourceRequestId);

                        TempData["SuccessMessage"] = $"Draft SPS Document '{model.DocumentNumber}' tersimpan. Bisa dilanjutkan kapan saja dari Approval Request Saya.";
                        return RedirectToAction("MyRequests", "Approval");
                    }

                    var allowed = await _approvalService.HasConsumableApprovalAsync(ApprovalActionType.SpsDocumentCreate, targetKey);
                    if (!allowed)
                    {
                        await _approvalService.CreateOrReusePendingRequestAsync(
                            ApprovalActionType.SpsDocumentCreate,
                            targetKey,
                            $"Request create SPS Document {model.DocumentNumber}",
                            Url.Action(nameof(Create), "SpsMaster") ?? "/SpsMaster/Create",
                            JsonSerializer.Serialize(model),
                            sourceRequestId);

                        TempData["SuccessMessage"] = $"Request data SPS Document '{model.DocumentNumber}' terkirim ke SUPERADMIN. Setelah disetujui, data akan otomatis dibuat.";
                        return RedirectToAction("MyRequests", "Approval");
                    }
                }

                // Check if DocumentNumber already exists
                if (await _context.SpsNoDocs.AnyAsync(s => s.DocumentNumber == model.DocumentNumber))
                {
                    TempData["ErrorMessage"] = "No. Document sudah ada di database";
                    ModelState.AddModelError("DocumentNumber", "No. Document sudah ada di database");
                    return View(model);
                }

                var spsNoDoc = SpsMapper.ToSpsNoDoc(model);
                
                if (!string.IsNullOrWhiteSpace(model.ItemList))
                {
                    var items = model.ItemList.Split(new[] { ',', '\n', '\r', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Select(i => i.Trim())
                                              .Where(i => !string.IsNullOrEmpty(i))
                                              .Distinct();
                    
                    foreach (var item in items)
                    {
                        spsNoDoc.ItemLists.Add(new SpsItemList 
                        { 
                            DocumentNumber = model.DocumentNumber, 
                            ItemList = item 
                        });
                    }
                }

                _context.SpsNoDocs.Add(spsNoDoc);
                await _context.SaveChangesAsync();

                await _approvalService.ConsumeApprovalAsync(ApprovalActionType.SpsDocumentCreate, model.DocumentNumber);

                TempData["SuccessMessage"] = $"âœ… Berhasil menambahkan Dokumen SPS '{model.DocumentNumber}'!";
                _logger.LogInformation($"Create Success: DocumentNumber '{model.DocumentNumber}' created");
                return RedirectToAction(nameof(Index));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error creating SPS Master");
                ModelState.AddModelError("", $"Error: {ex.Message}");
                return View(model);
            }
        }

        [HttpGet]
        public async Task<IActionResult> Edit(string? documentNumber, int? requestId)
        {
            if (string.IsNullOrWhiteSpace(documentNumber)) return NotFound();

            var spsNoDoc = await _context.SpsNoDocs
                .Include(s => s.ItemLists)
                .FirstOrDefaultAsync(s => s.DocumentNumber == documentNumber);

            if (spsNoDoc == null) return NotFound();

            // Convert to SpsMaster view model
            var model = SpsMapper.ToSpsMaster(spsNoDoc);

            if (requestId.HasValue)
            {
                var currentUser = HttpContext.Session.GetString("UserName") ?? string.Empty;
                var request = await _context.ApprovalRequests
                    .AsNoTracking()
                    .FirstOrDefaultAsync(r => r.Id == requestId.Value
                        && r.ActionType == ApprovalActionType.SpsDocumentEdit
                        && r.RequesterUserName.ToUpper() == currentUser.ToUpper()
                        && !string.IsNullOrWhiteSpace(r.PayloadJson));

                if (request != null)
                {
                    try
                    {
                        var modelFromRequest = JsonSerializer.Deserialize<SpsMaster>(request.PayloadJson!);
                        if (modelFromRequest != null)
                        {
                            model = modelFromRequest;
                            model.DocumentNumber = documentNumber;
                            ViewBag.SourceRequestId = request.Id;
                        }
                    }
                    catch
                    {
                        // Keep existing DB model if payload is not readable.
                    }
                }
            }

            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(string documentNumber, SpsMaster model, bool UpdateAllWithSameDocNumber = false, bool saveAsDraft = false, int? sourceRequestId = null)
        {
            if (documentNumber != model.DocumentNumber)
            {
                return NotFound();
            }

            // Auto-fill Revision Date (Format: dd-MM-yy)
            model.RevisionDate = DateTime.Now.ToString("dd-MM-yy");

            try
            {
                if (_approvalService.IsRequesterRole())
                {
                    if (saveAsDraft)
                    {
                        await _approvalService.SaveDraftRequestAsync(
                            ApprovalActionType.SpsDocumentEdit,
                            documentNumber,
                            $"Draft revisi SPS Document {documentNumber}",
                            Url.Action(nameof(Edit), "SpsMaster", new { documentNumber }) ?? $"/SpsMaster/Edit/{documentNumber}",
                            JsonSerializer.Serialize(model),
                            sourceRequestId);

                        TempData["SuccessMessage"] = $"Draft revisi dokumen '{documentNumber}' tersimpan. Bisa dilanjutkan kapan saja dari Approval Request Saya.";
                        return RedirectToAction("MyRequests", "Approval");
                    }

                    await _approvalService.CreateOrReusePendingRequestAsync(
                        ApprovalActionType.SpsDocumentEdit,
                        documentNumber,
                        $"Request revisi SPS Document {documentNumber}",
                        Url.Action(nameof(Edit), "SpsMaster", new { documentNumber }) ?? $"/SpsMaster/Edit/{documentNumber}",
                        JsonSerializer.Serialize(model),
                        sourceRequestId);

                    TempData["SuccessMessage"] = $"Request revisi dokumen '{documentNumber}' terkirim ke SUPERADMIN. Setelah disetujui, versi baru akan dibuat otomatis.";
                    return RedirectToAction("MyRequests", "Approval");
                }

                var spsNoDoc = await _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .FirstOrDefaultAsync(s => s.DocumentNumber == documentNumber);

                if (spsNoDoc == null) return NotFound();

                // Update SpsNoDoc properties
                SpsMapper.UpdateSpsNoDoc(spsNoDoc, model);

                // Update ItemLists
                if (spsNoDoc.ItemLists == null)
                {
                    spsNoDoc.ItemLists = new List<SpsItemList>();
                }
                
                spsNoDoc.ItemLists.Clear();
                
                if (!string.IsNullOrWhiteSpace(model.ItemList))
                {
                    var items = model.ItemList.Split(new[] { ',', '\n', '\r', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Select(i => i.Trim())
                                              .Where(i => !string.IsNullOrEmpty(i))
                                              .Distinct();
                    
                    foreach (var item in items)
                    {
                        spsNoDoc.ItemLists.Add(new SpsItemList 
                        { 
                            DocumentNumber = spsNoDoc.DocumentNumber, 
                            ItemList = item 
                        });
                    }
                }

                await _context.SaveChangesAsync();

                await _approvalService.ConsumeApprovalAsync(ApprovalActionType.SpsDocumentEdit, documentNumber);

                TempData["SuccessMessage"] = $"âœ… Update Berhasil: Dokumen '{documentNumber}' telah diperbarui.";
                _logger.LogInformation($"Edit SUCCESS: DocumentNumber '{documentNumber}' updated");

                return RedirectToAction(nameof(Index));
            }
            catch (DbUpdateConcurrencyException ex)
            {
                _logger.LogError(ex, $"Concurrency error updating '{documentNumber}'");
                if (!await SpsNoDocExists(documentNumber))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error updating SPS Master '{documentNumber}'");
                ModelState.AddModelError("", $"âŒ Error: {ex.Message}");
                return View(model);
            }
        }

        private async Task<bool> SpsNoDocExists(string documentNumber)
        {
            return await _context.SpsNoDocs.AnyAsync(e => e.DocumentNumber == documentNumber);
        }

        [HttpGet]
        public async Task<IActionResult> GetItemLists(string term)
        {
            if (string.IsNullOrWhiteSpace(term))
                return Json(new List<string>());

            var termUpper = term.ToUpper();
            var items = await _context.SpsItemLists
                .Where(i => i.ItemList != null && i.ItemList.ToUpper().Contains(termUpper))
                .Select(i => i.ItemList)
                .Distinct()
                .Take(20)
                .ToListAsync();

            return Json(items);
        }

        [HttpGet]
        public async Task<IActionResult> GetDocumentNumbers(string term)
        {
            if (string.IsNullOrWhiteSpace(term))
                return Json(new List<string>());

            var termUpper = term.ToUpper();
            var docs = await _context.SpsNoDocs
                .Where(s => s.IsActive && s.DocumentNumber != null && s.DocumentNumber.ToUpper().Contains(termUpper))
                .Select(s => s.DocumentNumber)
                .Distinct()
                .Take(20)
                .ToListAsync();

            return Json(docs);
        }

        [HttpGet]
        public async Task<IActionResult> SearchSps(string term)
        {
            if (string.IsNullOrWhiteSpace(term))
                return Json(new List<object>());

            var termUpper = term.ToUpper();
            
            // Search by ItemList
            var itemResults = await _context.SpsItemLists
                .Include(i => i.SpsNoDoc)
                .Where(i => i.ItemList != null && i.ItemList.ToUpper().Contains(termUpper))
                .Select(i => new {
                    id = i.DocumentNumber,
                    itemList = (string?)i.ItemList,
                    documentNumber = i.SpsNoDoc!.DocumentNumber,
                    hoseType = i.SpsNoDoc.HoseType,
                    customer = i.SpsNoDoc.Customer,
                    type = "ItemList"
                })
                .Take(10)
                .ToListAsync();

            // Search by DocumentNumber or HoseType
            var docResults = await _context.SpsNoDocs
                .Where(s => 
                    (s.DocumentNumber != null && s.DocumentNumber.ToUpper().Contains(termUpper)) ||
                    (s.HoseType != null && s.HoseType.ToUpper().Contains(termUpper)))
                .Select(s => new {
                    id = s.DocumentNumber,
                    itemList = (string?)null,
                    documentNumber = s.DocumentNumber,
                    hoseType = s.HoseType,
                    customer = s.Customer,
                    type = "Document"
                })
                .Take(10)
                .ToListAsync();

            var combined = itemResults.Concat(docResults)
                .DistinctBy(x => x.id)
                .Take(20)
                .ToList();

            return Json(combined);
        }

        [HttpPost]
        public async Task<IActionResult> DeleteMultiple([FromBody] List<string> documentNumbers)
        {
            try
            {
                if (documentNumbers == null || !documentNumbers.Any())
                {
                    return Json(new { success = false, message = "Tidak ada Document Number yang dipilih" });
                }

                var recordsToDelete = await _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .Where(s => documentNumbers.Contains(s.DocumentNumber))
                    .ToListAsync();

                if (!recordsToDelete.Any())
                {
                    return Json(new { success = false, message = "Data tidak ditemukan" });
                }

                _context.SpsNoDocs.RemoveRange(recordsToDelete);
                await _context.SaveChangesAsync();

                _logger.LogInformation($"Berhasil menghapus {recordsToDelete.Count} SPS Master records");

                return Json(new { 
                    success = true, 
                    message = $"Berhasil menghapus {recordsToDelete.Count} data SPS Master",
                    deletedCount = recordsToDelete.Count
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saat menghapus multiple SPS Master");
                return Json(new { 
                    success = false, 
                    message = $"Terjadi kesalahan: {ex.Message}" 
                });
            }
        }


        [HttpGet]
        public IActionResult DownloadTemplateFull()
        {
            try
            {
                using var package = new ExcelPackage();
                var ws = package.Workbook.Worksheets.Add("SPS");

                // Helper: tulis group label (row 1, merged)
                void SetGroup(int c1, int c2, string label, System.Drawing.Color bg)
                {
                    ws.Cells[1, c1, 1, c2].Merge = true;
                    ws.Cells[1, c1].Value = label;
                    using var r = ws.Cells[1, c1, 1, c2];
                    r.Style.Font.Bold = true;
                    r.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(bg);
                    r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    r.Style.VerticalAlignment   = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                }

                // Definisi kolom: (header, groupIndex)
                // Group: 0=Identifikasi, 1=Dokumen, 2=Material, 3=Dies, 4=Yarn,
                //        5=Suhu, 6=Speed/Feed, 7=Toleransi, 8=QM, 9=ItemList
                var colDefs = new (string H, int G)[] {
                    // 0 IDENTIFIKASI
                    ("ID Excel",0),("No",0),("MC",0),
                    // 1 DOKUMEN & PRODUK
                    ("NO. DOC",1),("NO. REV",1),("REV. DATE",1),("CUSTOMER",1),
                    ("HOSE TYPE",1),("DIMENSI",1),("FORMULASI",1),("MACHINE",1),
                    // 2 MATERIAL
                    ("MATERIAL TSM",2),("INNER | #tube",2),("OUTER | #cover",2),
                    ("MIDDLE | #tube2",2),("USE LIMITS INNER",2),("USE LIMITS OUTER",2),("USE LIMITS MIDDLE",2),
                    // 3 DIES & TOOLS â€” setiap dies: Raw | MIN | STD | MAX
                    ("NIPPLE",3),("NIPPLE MIN",3),("NIPPLE STD",3),("NIPPLE MAX",3),
                    ("TUBE DIE",3),("TUBE DIE MIN",3),("TUBE DIE STD",3),("TUBE DIE MAX",3),
                    ("COVER DIE",3),("COVER DIE MIN",3),("COVER DIE STD",3),("COVER DIE MAX",3),
                    ("MIDDLE DIE",3),("MIDDLE DIE MIN",3),("MIDDLE DIE STD",3),("MIDDLE DIE MAX",3),
                    ("SPACER DIE",3),("SPACER DIE MIN",3),("SPACER DIE STD",3),("SPACER DIE MAX",3),
                    ("A DISTANCE",3),("A DISTANCE MIN",3),("A DISTANCE STD",3),("A DISTANCE MAX",3),
                    // 4 YARN & MESH
                    ("YARN",4),("TENSION YARN INNER",4),("TENSION YARN OUTER",4),
                    ("MESH SCREEN 1",4),("MESH DIM 1",4),("MESH DIM 1 MIN",4),("MESH DIM 1 STD",4),("MESH DIM 1 MAX",4),
                    ("MESH SCREEN 2",4),("MESH DIM 2",4),("MESH DIM 2 MIN",4),("MESH DIM 2 STD",4),("MESH DIM 2 MAX",4),
                    ("MESH SCREEN 3",4),("MESH DIM 3",4),("MESH DIM 3 MIN",4),("MESH DIM 3 STD",4),("MESH DIM 3 MAX",4),
                    ("PITCH YARN",4),("PITCH YARN MIN",4),("PITCH YARN STD",4),("PITCH YARN MAX",4),
                    // 5 SUHU & TEKANAN â€” Mat1
                    ("HEAD TEMP 1",5),("HEAD TEMP 1 MIN",5),("HEAD TEMP 1 STD",5),("HEAD TEMP 1 MAX",5),
                    ("CYL 1-1",5),("CYL 1-1 MIN",5),("CYL 1-1 STD",5),("CYL 1-1 MAX",5),
                    ("CYL 2-1",5),("CYL 2-1 MIN",5),("CYL 2-1 STD",5),("CYL 2-1 MAX",5),
                    ("CYL 3-1",5),("CYL 3-1 MIN",5),("CYL 3-1 STD",5),("CYL 3-1 MAX",5),
                    ("SCREW TEMP 1",5),("SCREW TEMP 1 MIN",5),("SCREW TEMP 1 STD",5),("SCREW TEMP 1 MAX",5),
                    ("SCREW SPEED 1",5),("SCREW SPEED 1 MIN",5),("SCREW SPEED 1 STD",5),("SCREW SPEED 1 MAX",5),
                    ("PRESSURE 1",5),("PRESSURE 1 MIN",5),("PRESSURE 1 STD",5),("PRESSURE 1 MAX",5),
                    // Mat2
                    ("HEAD TEMP 2",5),("HEAD TEMP 2 MIN",5),("HEAD TEMP 2 STD",5),("HEAD TEMP 2 MAX",5),
                    ("CYL 1-2",5),("CYL 1-2 MIN",5),("CYL 1-2 STD",5),("CYL 1-2 MAX",5),
                    ("CYL 2-2",5),("CYL 2-2 MIN",5),("CYL 2-2 STD",5),("CYL 2-2 MAX",5),
                    ("CYL 3-2",5),("CYL 3-2 MIN",5),("CYL 3-2 STD",5),("CYL 3-2 MAX",5),
                    ("SCREW TEMP 2",5),("SCREW TEMP 2 MIN",5),("SCREW TEMP 2 STD",5),("SCREW TEMP 2 MAX",5),
                    ("SCREW SPEED 2",5),("SCREW SPEED 2 MIN",5),("SCREW SPEED 2 STD",5),("SCREW SPEED 2 MAX",5),
                    ("PRESSURE 2",5),("PRESSURE 2 MIN",5),("PRESSURE 2 STD",5),("PRESSURE 2 MAX",5),
                    // Mat3
                    ("HEAD TEMP 3",5),("HEAD TEMP 3 MIN",5),("HEAD TEMP 3 STD",5),("HEAD TEMP 3 MAX",5),
                    ("CYL 1-3",5),("CYL 1-3 MIN",5),("CYL 1-3 STD",5),("CYL 1-3 MAX",5),
                    ("CYL 2-3",5),("CYL 2-3 MIN",5),("CYL 2-3 STD",5),("CYL 2-3 MAX",5),
                    ("CYL 3-3",5),("CYL 3-3 MIN",5),("CYL 3-3 STD",5),("CYL 3-3 MAX",5),
                    ("SCREW TEMP 3",5),("SCREW TEMP 3 MIN",5),("SCREW TEMP 3 STD",5),("SCREW TEMP 3 MAX",5),
                    ("SCREW SPEED 3",5),("SCREW SPEED 3 MIN",5),("SCREW SPEED 3 STD",5),("SCREW SPEED 3 MAX",5),
                    ("PRESSURE 3",5),("PRESSURE 3 MIN",5),("PRESSURE 3 STD",5),("PRESSURE 3 MAX",5),
                    // 6 SPEED / FEED / OUTPUT
                    ("CURRENT VALUE",6),
                    ("FEED 1",6),("FEED 1 MIN",6),("FEED 1 STD",6),("FEED 1 MAX",6),
                    ("FEED 2",6),("FEED 2 MIN",6),("FEED 2 STD",6),("FEED 2 MAX",6),
                    ("FEED 3",6),("FEED 3 MIN",6),("FEED 3 STD",6),("FEED 3 MAX",6),
                    ("FEED ROLL RATIO 1",6),("F.R.RATIO 1 MIN",6),("F.R.RATIO 1 STD",6),("F.R.RATIO 1 MAX",6),
                    ("FEED ROLL RATIO 2",6),("F.R.RATIO 2 MIN",6),("F.R.RATIO 2 STD",6),("F.R.RATIO 2 MAX",6),
                    ("FEED ROLL RATIO 3",6),("F.R.RATIO 3 MIN",6),("F.R.RATIO 3 STD",6),("F.R.RATIO 3 MAX",6),
                    ("AM METER 1",6),("AM MTR 1 MIN",6),("AM MTR 1 STD",6),("AM MTR 1 MAX",6),
                    ("AM METER 2",6),("AM MTR 2 MIN",6),("AM MTR 2 STD",6),("AM MTR 2 MAX",6),
                    ("AM METER 3",6),("AM MTR 3 MIN",6),("AM MTR 3 STD",6),("AM MTR 3 MAX",6),
                    ("PRESET VALUE",6),("PRESET MIN",6),("PRESET STD",6),("PRESET MAX",6),
                    ("CONTROL VALUE",6),("CTRL MIN",6),("CTRL STD",6),("CTRL MAX",6),
                    ("SPIRAL PITCH SETTING",6),("SPS MIN",6),("SPS STD",6),("SPS MAX",6),
                    ("SPIRAL PITCH DISPLAY",6),("SPD MIN",6),("SPD STD",6),("SPD MAX",6),
                    ("SPIRAL SPEED",6),("SPSPD MIN",6),("SPSPD STD",6),("SPSPD MAX",6),
                    ("HOSE SPEED",6),("HSPD MIN",6),("HSPD STD",6),("HSPD MAX",6),
                    ("CHILLER WATER TEMP",6),("CHILL MIN",6),("CHILL STD",6),("CHILL MAX",6),
                    ("DANCER POSITION",6),("DANCER MIN",6),("DANCER STD",6),("DANCER MAX",6),
                    ("CATERPILLAR GAP",6),("CAT MIN",6),("CAT STD",6),("CAT MAX",6),
                    ("CUTTING SPEED",6),("CUT MIN",6),("CUT STD",6),("CUT MAX",6),
                    ("TAKE UP CONV SPD",6),("TAKU MIN",6),("TAKU STD",6),("TAKU MAX",6),
                    ("COOL CONV SPD 1",6),("CC1 MIN",6),("CC1 STD",6),("CC1 MAX",6),
                    ("COOL CONV SPD 2",6),("CC2 MIN",6),("CC2 STD",6),("CC2 MAX",6),
                    ("CONVEYOR RATIO",6),("CONV R MIN",6),("CONV R STD",6),("CONV R MAX",6),
                    ("OD SENSOR",6),("OD MIN",6),("OD STD",6),("OD MAX",6),
                    ("MARKING SORT",6),("TEXT MARKING MT'L",6),("MARKING COLOUR",6),("UNSMOOTH SURFACE",6),
                    // 7 TOLERANSI & TEBAL
                    ("Â± INNER",7),("TOL INNER MIN",7),("TOL INNER STD",7),("TOL INNER MAX",7),
                    ("Â± OUTER",7),("TOL OUTER MIN",7),("TOL OUTER STD",7),("TOL OUTER MAX",7),
                    ("TEBAL INNER",7),("TB INNER MIN",7),("TB INNER STD",7),("TB INNER MAX",7),
                    ("TEBAL OUTER",7),("TB OUTER MIN",7),("TB OUTER STD",7),("TB OUTER MAX",7),
                    ("TEBAL INNER+MID",7),("TB I+M MIN",7),("TB I+M STD",7),("TB I+M MAX",7),
                    ("TEBAL TOTAL",7),("TB TOTAL MIN",7),("TB TOTAL STD",7),("TB TOTAL MAX",7),
                    ("SELISIH TEBAL",7),("SELISIH MIN",7),("SELISIH STD",7),("SELISIH MAX",7),
                    // 8 QUALITY MATRIX
                    ("INNER TARGET",8),("INNER TOL",8),("INNER LCL",8),("INNER MIN",8),("INNER UCL",8),("INNER MAX",8),
                    ("I+MID TARGET",8),("I+MID TOL",8),("I+MID LCL",8),("I+MID MIN",8),("I+MID UCL",8),("I+MID MAX",8),
                    ("THICK TARGET",8),("THICK TOL",8),("THICK LCL",8),("THICK MIN",8),("THICK UCL",8),("THICK MAX",8),
                    ("TOTAL TARGET",8),("TOTAL TOL",8),("TOTAL LCL",8),("TOTAL MIN",8),("TOTAL UCL",8),("TOTAL MAX",8),
                };

                int totalCols = colDefs.Length;

                var groupColors = new System.Drawing.Color[] {
                    System.Drawing.Color.FromArgb(31, 78, 121),   // 0 Identifikasi
                    System.Drawing.Color.FromArgb(21, 96, 130),   // 1 Dokumen
                    System.Drawing.Color.FromArgb(0, 97, 0),      // 2 Material
                    System.Drawing.Color.FromArgb(166, 74, 0),    // 3 Dies
                    System.Drawing.Color.FromArgb(0, 83, 83),     // 4 Yarn
                    System.Drawing.Color.FromArgb(148, 17, 0),    // 5 Suhu
                    System.Drawing.Color.FromArgb(12, 80, 20),    // 6 Speed/Feed
                    System.Drawing.Color.FromArgb(120, 0, 0),     // 7 Toleransi
                    System.Drawing.Color.FromArgb(0, 60, 80),     // 8 QM
                };

                var groupNames = new string[] {
                    "IDENTIFIKASI",
                    "DOKUMEN & PRODUK",
                    "MATERIAL",
                    "DIES & TOOLS (format X±Y, atau isi MIN/STD/MAX manual)",
                    "YARN & MESH",
                    "SUHU, SCREW & TEKANAN (format X±Y, atau isi MIN/STD/MAX manual)",
                    "SPEED, FEED & OUTPUT MESIN (format X±Y, atau isi MIN/STD/MAX manual)",
                    "TOLERANSI & TEBAL (format X±Y, atau isi MIN/STD/MAX manual)",
                    "QUALITY MATRIX",
                };

                // Tulis header kolom (Row 2)
                for (int i = 0; i < totalCols; i++)
                {
                    var cell = ws.Cells[2, i + 1];
                    cell.Value = colDefs[i].H;
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(groupColors[colDefs[i].G]);
                    cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment   = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    cell.Style.WrapText = true;
                    cell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    cell.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.White);
                }

                // Tulis group label Row 1 — deteksi batas grup
                int gStart = 1;
                int curG   = colDefs[0].G;
                for (int i = 1; i <= totalCols; i++)
                {
                    int g = (i < totalCols) ? colDefs[i].G : -1;
                    if (g != curG)
                    {
                        SetGroup(gStart, i, groupNames[curG], groupColors[curG]);
                        gStart = i + 1;
                        curG   = g;
                    }
                }

                // Bikin sheet terpisah untuk petunjuk
                var wsPetunjuk = package.Workbook.Worksheets.Add("📋 Petunjuk");
                var notes = new string[] {
                    "PETUNJUK PENGISIAN TEMPLATE SPS:",
                    "1. Kolom MIN/STD/MAX: isi nilai numerik langsung (MIN=batas bawah, STD=standar, MAX=batas atas).",
                    "2. Kolom ± (contoh: ± INNER): isi format X±Y (contoh: 37.0±0.2). Sistem otomatis hitung MIN & MAX.",
                    "3. Jika keduanya diisi, sistem menyimpan keduanya — nilai manual & hitungan otomatis sebagai pembanding.",
                    "4. Item List dikelola secara terpisah melalui menu Master Item List (dengan Length Shrink & Length Cutting).",
                    "5. Sheet name untuk data utama harus: 'SPS', 'Digitalisasi', 'Parameter Setting', atau 'Master'.",
                    "6. Data diisi di sheet SPS mulai BARIS 3 ke bawah (langsung di bawah header tabel)."
                };
                for (int n = 0; n < notes.Length; n++)
                {
                    var nc = wsPetunjuk.Cells[2 + n, 2]; // B2 ke bawah
                    nc.Value = notes[n];
                    nc.Style.Font.Bold   = (n == 0);
                    nc.Style.Font.Size   = (n == 0) ? 12 : 11;
                    nc.Style.Font.Italic = (n > 0);
                    nc.Style.Font.Color.SetColor(n == 0
                        ? System.Drawing.Color.Red
                        : System.Drawing.Color.FromArgb(50, 50, 50));
                }
                wsPetunjuk.Cells.AutoFitColumns();
                ws.View.TabSelected = true;

                // Freeze & sizing
                ws.View.FreezePanes(3, 1);
                ws.Cells.AutoFitColumns();
                for (int col = 1; col <= totalCols; col++)
                {
                    if (ws.Column(col).Width < 8)  ws.Column(col).Width = 8;
                    if (ws.Column(col).Width > 28) ws.Column(col).Width = 28;
                }
                ws.Row(1).Height = 24;
                ws.Row(2).Height = 45;

                var stream   = new MemoryStream(package.GetAsByteArray());
                var fileName = $"Template_SPS_{DateTime.Now:yyyyMMdd}.xlsx";
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating SPS template");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        [HttpGet]
        public async Task<IActionResult> ExportExcel()
        {
            try
            {
                var allSpsNoDocs = await _context.SpsNoDocs
                    .OrderBy(s => s.DocumentNumber)
                    .ToListAsync();

                // 1 Record per SPS Document
                var listData = allSpsNoDocs.Select(nodoc => SpsMapper.ToSpsMaster(nodoc, "")).ToList();

                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("SPS Master Data");

                // Definisi kolom persis seperti UI (Index.cshtml)
                var colDefs = new List<(string H, Func<SpsMaster, object?> Value, int G, string Color)> {
                    ("ID Excel", x => x.No, 0, ""),
                    ("NO", x => x.No, 0, ""),
                    ("Machine", x => x.Machine, 0, ""),
                    ("NO. DOC", x => x.DocumentNumber, 0, ""),
                    ("STATUS", x => x.IsActive ? "ACTIVE" : "INACTIVE", 0, ""),
                    ("NO. REV", x => x.RevisionNumber, 0, ""),
                    ("REV. DATE", x => x.RevisionDate, 0, ""),
                    ("FORMULASI", x => x.Formulasi, 0, ""),
                    ("MC", x => x.MachineCode, 0, ""),

                    ("Customer", x => x.Customer, 1, ""),
                    ("Hose Type", x => x.HoseType, 1, ""),
                    ("Dimensi", x => x.Dimensi, 1, ""),
                    ("Material", x => x.Material, 1, ""),
                    ("Inner Tube", x => x.InnerTube, 1, ""),
                    ("Middle Tube", x => x.MiddleTube, 1, ""),
                    ("Outer Cover", x => x.OuterCover, 1, ""),
                    ("Use Limits In", x => x.UseLimitsInner, 1, ""),
                    ("Use Limits Mid", x => x.UseLimitsMiddle, 1, ""),
                    ("Use Limits Out", x => x.UseLimitsOuter, 1, ""),
                    ("Yarn", x => x.Yarn, 1, ""),
                    ("MIN Pitch Yarn", x => x.PitchYarn_Min?.ToString("F2"), 1, "success"),
                    ("STD Pitch Yarn", x => x.PitchYarn_Asli?.ToString("F2"), 1, "primary"),
                    ("MAX Pitch Yarn", x => x.PitchYarn_Max?.ToString("F2"), 1, "danger"),
                    ("Tension In", x => x.TensionYarnInner, 1, ""),
                    ("Tension Out", x => x.TensionYarnOuter, 1, ""),
                    ("MIN Nipple", x => x.Nipple_Min?.ToString("F2"), 1, "success"),
                    ("STD Nipple", x => x.Nipple_Asli?.ToString("F2"), 1, "primary"),
                    ("MAX Nipple", x => x.Nipple_Max?.ToString("F2"), 1, "danger"),
                    ("MIN Tube Die", x => x.TubeDie_Min?.ToString("F2"), 1, "success"),
                    ("STD Tube Die", x => x.TubeDie_Asli?.ToString("F2"), 1, "primary"),
                    ("MAX Tube Die", x => x.TubeDie_Max?.ToString("F2"), 1, "danger"),
                    ("MIN Middle Die", x => x.MiddleDie_Min?.ToString("F2"), 1, "success"),
                    ("STD Middle Die", x => x.MiddleDie_Asli?.ToString("F2"), 1, "primary"),
                    ("MAX Middle Die", x => x.MiddleDie_Max?.ToString("F2"), 1, "danger"),
                    ("MIN Cover Die", x => x.CoverDie_Min?.ToString("F2"), 1, "success"),
                    ("STD Cover Die", x => x.CoverDie_Asli?.ToString("F2"), 1, "primary"),
                    ("MAX Cover Die", x => x.CoverDie_Max?.ToString("F2"), 1, "danger"),
                    ("MIN Spacer", x => x.SpacerDie_Min?.ToString("F2"), 1, "success"),
                    ("STD Spacer", x => x.SpacerDie_Asli?.ToString("F2"), 1, "primary"),
                    ("MAX Spacer", x => x.SpacerDie_Max?.ToString("F2"), 1, "danger"),
                    ("MIN A Distance", x => x.ADistance_Min?.ToString("F2"), 1, "success"),
                    ("STD A Distance", x => x.ADistance_Asli?.ToString("F2"), 1, "primary"),
                    ("MAX A Distance", x => x.ADistance_Max?.ToString("F2"), 1, "danger"),

                    ("Mesh Screen 1", x => x.MeshScreen1, 2, ""),
                    ("MIN Mesh Dim 1", x => x.MeshDim1_Min?.ToString("F2"), 2, "success"),
                    ("STD Mesh Dim 1", x => x.MeshDim1_Asli?.ToString("F2"), 2, "primary"),
                    ("MAX Mesh Dim 1", x => x.MeshDim1_Max?.ToString("F2"), 2, "danger"),
                    ("Mesh Screen 2", x => x.MeshScreen2, 2, ""),
                    ("MIN Mesh Dim 2", x => x.MeshDim2_Min?.ToString("F2"), 2, "success"),
                    ("STD Mesh Dim 2", x => x.MeshDim2_Asli?.ToString("F2"), 2, "primary"),
                    ("MAX Mesh Dim 2", x => x.MeshDim2_Max?.ToString("F2"), 2, "danger"),
                    ("Mesh Screen 3", x => x.MeshScreen3, 2, ""),
                    ("MIN Mesh Dim 3", x => x.MeshDim3_Min?.ToString("F2"), 2, "success"),
                    ("STD Mesh Dim 3", x => x.MeshDim3_Asli?.ToString("F2"), 2, "primary"),
                    ("MAX Mesh Dim 3", x => x.MeshDim3_Max?.ToString("F2"), 2, "danger"),

                    ("MIN Head 1", x => x.HeadTemp1_Min?.ToString("F2"), 3, "success"),
                    ("STD Head 1", x => x.HeadTemp1_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Head 1", x => x.HeadTemp1_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 1-1", x => x.Cylinder1_1_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 1-1", x => x.Cylinder1_1_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 1-1", x => x.Cylinder1_1_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 2-1", x => x.Cylinder2_1_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 2-1", x => x.Cylinder2_1_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 2-1", x => x.Cylinder2_1_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 3-1", x => x.Cylinder3_1_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 3-1", x => x.Cylinder3_1_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 3-1", x => x.Cylinder3_1_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Screw 1", x => x.ScrewTemp1_Min?.ToString("F2"), 3, "success"),
                    ("STD Screw 1", x => x.ScrewTemp1_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Screw 1", x => x.ScrewTemp1_Max?.ToString("F2"), 3, "danger"),

                    ("MIN Head 2", x => x.HeadTemp2_Min?.ToString("F2"), 3, "success"),
                    ("STD Head 2", x => x.HeadTemp2_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Head 2", x => x.HeadTemp2_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 1-2", x => x.Cylinder1_2_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 1-2", x => x.Cylinder1_2_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 1-2", x => x.Cylinder1_2_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 2-2", x => x.Cylinder2_2_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 2-2", x => x.Cylinder2_2_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 2-2", x => x.Cylinder2_2_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 3-2", x => x.Cylinder3_2_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 3-2", x => x.Cylinder3_2_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 3-2", x => x.Cylinder3_2_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Screw 2", x => x.ScrewTemp2_Min?.ToString("F2"), 3, "success"),
                    ("STD Screw 2", x => x.ScrewTemp2_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Screw 2", x => x.ScrewTemp2_Max?.ToString("F2"), 3, "danger"),

                    ("MIN Head 3", x => x.HeadTemp3_Min?.ToString("F2"), 3, "success"),
                    ("STD Head 3", x => x.HeadTemp3_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Head 3", x => x.HeadTemp3_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 1-3", x => x.Cylinder1_3_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 1-3", x => x.Cylinder1_3_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 1-3", x => x.Cylinder1_3_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 2-3", x => x.Cylinder2_3_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 2-3", x => x.Cylinder2_3_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 2-3", x => x.Cylinder2_3_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Cyl 3-3", x => x.Cylinder3_3_Min?.ToString("F2"), 3, "success"),
                    ("STD Cyl 3-3", x => x.Cylinder3_3_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Cyl 3-3", x => x.Cylinder3_3_Max?.ToString("F2"), 3, "danger"),
                    ("MIN Screw 3", x => x.ScrewTemp3_Min?.ToString("F2"), 3, "success"),
                    ("STD Screw 3", x => x.ScrewTemp3_Asli?.ToString("F2"), 3, "primary"),
                    ("MAX Screw 3", x => x.ScrewTemp3_Max?.ToString("F2"), 3, "danger"),

                    ("MIN Feed 1", x => x.Feed1_Min?.ToString("F2"), 4, "success"),
                    ("STD Feed 1", x => x.Feed1_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Feed 1", x => x.Feed1_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Feed 2", x => x.Feed2_Min?.ToString("F2"), 4, "success"),
                    ("STD Feed 2", x => x.Feed2_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Feed 2", x => x.Feed2_Max?.ToString("F2"), 4, "danger"),

                    ("MIN Screw V1", x => x.ScrewSpeed1_Min?.ToString("F2"), 4, "success"),
                    ("STD Screw V1", x => x.ScrewSpeed1_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Screw V1", x => x.ScrewSpeed1_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Screw V2", x => x.ScrewSpeed2_Min?.ToString("F2"), 4, "success"),
                    ("STD Screw V2", x => x.ScrewSpeed2_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Screw V2", x => x.ScrewSpeed2_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Screw V3", x => x.ScrewSpeed3_Min?.ToString("F2"), 4, "success"),
                    ("STD Screw V3", x => x.ScrewSpeed3_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Screw V3", x => x.ScrewSpeed3_Max?.ToString("F2"), 4, "danger"),

                    ("MIN Feed Roll 1", x => x.FeedRollRatio1_Min?.ToString("F2"), 4, "success"),
                    ("STD Feed Roll 1", x => x.FeedRollRatio1_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Feed Roll 1", x => x.FeedRollRatio1_Asli == null && !string.IsNullOrEmpty(x.FeedRollRatio1) ? x.FeedRollRatio1 : x.FeedRollRatio1_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Feed Roll 2", x => x.FeedRollRatio2_Min?.ToString("F2"), 4, "success"),
                    ("STD Feed Roll 2", x => x.FeedRollRatio2_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Feed Roll 2", x => x.FeedRollRatio2_Asli == null && !string.IsNullOrEmpty(x.FeedRollRatio2) ? x.FeedRollRatio2 : x.FeedRollRatio2_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Feed 3", x => x.Feed3_Min?.ToString("F2"), 4, "success"),
                    ("STD Feed 3", x => x.Feed3_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Feed 3", x => x.Feed3_Asli == null && !string.IsNullOrEmpty(x.FeedRollRatio3) ? x.FeedRollRatio3 : x.Feed3_Max?.ToString("F2"), 4, "danger"),

                    ("MIN Press 1", x => x.Pressure1_Min?.ToString("F2"), 4, "success"),
                    ("STD Press 1", x => x.Pressure1_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Press 1", x => x.Pressure1_Asli == null && !string.IsNullOrEmpty(x.Pressure1) ? x.Pressure1 : x.Pressure1_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Press 2", x => x.Pressure2_Min?.ToString("F2"), 4, "success"),
                    ("STD Press 2", x => x.Pressure2_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Press 2", x => x.Pressure2_Asli == null && !string.IsNullOrEmpty(x.Pressure2) ? x.Pressure2 : x.Pressure2_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Press 3", x => x.Pressure3_Min?.ToString("F2"), 4, "success"),
                    ("STD Press 3", x => x.Pressure3_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Press 3", x => x.Pressure3_Asli == null && !string.IsNullOrEmpty(x.Pressure3) ? x.Pressure3 : x.Pressure3_Max?.ToString("F2"), 4, "danger"),

                    ("Curr Val", x => x.CurrentValue, 4, ""),
                    ("MIN Am 1", x => x.AmMeter_Min?.ToString("F2"), 4, "success"),
                    ("STD Am 1", x => x.AmMeter_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Am 1", x => x.AmMeter_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Am 2", x => x.AmMeter2_Min?.ToString("F2"), 4, "success"),
                    ("STD Am 2", x => x.AmMeter2_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Am 2", x => x.AmMeter2_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Am 3", x => x.AmMeter3_Min?.ToString("F2"), 4, "success"),
                    ("STD Am 3", x => x.AmMeter3_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Am 3", x => x.AmMeter3_Max?.ToString("F2"), 4, "danger"),

                    ("MIN Preset", x => x.PresetValue_Min?.ToString("F2"), 4, "success"),
                    ("STD Preset", x => x.PresetValue_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Preset", x => x.PresetValue_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Control", x => x.ControlValue_Min?.ToString("F2"), 4, "success"),
                    ("STD Control", x => x.ControlValue_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Control", x => x.ControlValue_Max?.ToString("F2"), 4, "danger"),

                    ("MIN Sp-Set", x => x.SpiralPitchSetting_Min?.ToString("F2"), 4, "success"),
                    ("STD Sp-Set", x => x.SpiralPitchSetting_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Sp-Set", x => x.SpiralPitchSetting_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Sp-Disp", x => x.SpiralPitchDisplay_Min?.ToString("F2"), 4, "success"),
                    ("STD Sp-Disp", x => x.SpiralPitchDisplay_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Sp-Disp", x => x.SpiralPitchDisplay_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Spiral Speed", x => x.SpiralSpeed_Min?.ToString("F2"), 4, "success"),
                    ("STD Spiral Speed", x => x.SpiralSpeed_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Spiral Speed", x => x.SpiralSpeed_Max?.ToString("F2"), 4, "danger"),
                    ("MIN Hose Speed", x => x.HoseSpeed_Min?.ToString("F2"), 4, "success"),
                    ("STD Hose Speed", x => x.HoseSpeed_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX Hose Speed", x => x.HoseSpeed_Max?.ToString("F2"), 4, "danger"),
                    ("Unsmooth", x => x.UnsmoothSurface, 4, ""),
                    ("MIN OD Sensor", x => x.OdSensor_Min?.ToString("F2"), 4, "success"),
                    ("STD OD Sensor", x => x.OdSensor_Asli?.ToString("F2"), 4, "primary"),
                    ("MAX OD Sensor", x => x.OdSensor_Max?.ToString("F2"), 4, "danger"),

                    ("MIN Chiller", x => x.ChillerWaterTemp_Min?.ToString("F2"), 5, "success"),
                    ("STD Chiller", x => x.ChillerWaterTemp_Asli?.ToString("F2"), 5, "primary"),
                    ("MAX Chiller", x => x.ChillerWaterTemp_Asli == null && !string.IsNullOrEmpty(x.ChillerWaterTemp) ? x.ChillerWaterTemp : x.ChillerWaterTemp_Max?.ToString("F2"), 5, "danger"),
                    ("MIN Dancer", x => x.DancerPosition_Min?.ToString("F2"), 5, "success"),
                    ("STD Dancer", x => x.DancerPosition_Asli?.ToString("F2"), 5, "primary"),
                    ("MAX Dancer", x => x.DancerPosition_Max?.ToString("F2"), 5, "danger"),
                    ("MIN Caterpillar Gap", x => x.CaterpillarGap_Min?.ToString("F2"), 5, "success"),
                    ("STD Caterpillar Gap", x => x.CaterpillarGap_Asli?.ToString("F2"), 5, "primary"),
                    ("MAX Caterpillar Gap", x => x.CaterpillarGap_Max?.ToString("F2"), 5, "danger"),
                    ("MIN Cutting", x => x.CuttingSpeed_Min?.ToString("F2"), 5, "success"),
                    ("STD Cutting", x => x.CuttingSpeed_Asli?.ToString("F2"), 5, "primary"),
                    ("MAX Cutting", x => x.CuttingSpeed_Max?.ToString("F2"), 5, "danger"),
                    ("MIN Take-up Speed", x => x.TakeUpConveyorSpeed_Min?.ToString("F2"), 5, "success"),
                    ("STD Take-up Speed", x => x.TakeUpConveyorSpeed_Asli?.ToString("F2"), 5, "primary"),
                    ("MAX Take-up Speed", x => x.TakeUpConveyorSpeed_Max?.ToString("F2"), 5, "danger"),
                    ("MIN Cool-S 1", x => x.CoolConveyorSpeed_Min?.ToString("F2"), 5, "success"),
                    ("STD Cool-S 1", x => x.CoolConveyorSpeed_Asli?.ToString("F2"), 5, "primary"),
                    ("MAX Cool-S 1", x => x.CoolConveyorSpeed_Max?.ToString("F2"), 5, "danger"),
                    ("MIN Cool-S 2", x => x.CoolConveyorSpeed2_Min?.ToString("F2"), 5, "success"),
                    ("STD Cool-S 2", x => x.CoolConveyorSpeed2_Asli?.ToString("F2"), 5, "primary"),
                    ("MAX Cool-S 2", x => x.CoolConveyorSpeed2_Max?.ToString("F2"), 5, "danger"),
                    ("MIN Conv-R", x => x.ConveyorRatio_Min?.ToString("F2"), 5, "success"),
                    ("STD Conv-R", x => x.ConveyorRatio_Asli?.ToString("F2"), 5, "primary"),
                    ("MAX Conv-R", x => x.ConveyorRatio_Max?.ToString("F2"), 5, "danger"),
                    ("Mark Sort", x => x.MarkingSort, 5, ""),
                    ("Text Material", x => x.TextMarkingMaterial, 5, ""),
                    ("Mark Colour", x => x.MarkingColour, 5, ""),

                    ("MIN ± Inner", x => x.ToleranceInner_Min?.ToString("F2"), 6, "success"),
                    ("STD ± Inner", x => x.ToleranceInner_Asli?.ToString("F2"), 6, "primary"),
                    ("MAX ± Inner", x => x.ToleranceInner_Max?.ToString("F2"), 6, "danger"),
                    ("MIN ± Outer", x => x.ToleranceOuter_Min?.ToString("F2"), 6, "success"),
                    ("STD ± Outer", x => x.ToleranceOuter_Asli?.ToString("F2"), 6, "primary"),
                    ("MAX ± Outer", x => x.ToleranceOuter_Max?.ToString("F2"), 6, "danger"),
                    ("MIN Tebal Inner", x => x.TebalInner_Min?.ToString("F2"), 6, "success"),
                    ("STD Tebal Inner", x => x.TebalInner_Asli?.ToString("F2"), 6, "primary"),
                    ("MAX Tebal Inner", x => x.TebalInner_Max?.ToString("F2"), 6, "danger"),
                    ("MIN Inner+Middle", x => x.TebalInnerMiddle_Min?.ToString("F2"), 6, "success"),
                    ("STD Inner+Middle", x => x.TebalInnerMiddle_Asli?.ToString("F2"), 6, "primary"),
                    ("MAX Inner+Middle", x => x.TebalInnerMiddle_Max?.ToString("F2"), 6, "danger"),
                    ("MIN Tebal Outer", x => x.TebalOuter_Min?.ToString("F2"), 6, "success"),
                    ("STD Tebal Outer", x => x.TebalOuter_Asli?.ToString("F2"), 6, "primary"),
                    ("MAX Tebal Outer", x => x.TebalOuter_Max?.ToString("F2"), 6, "danger"),
                    ("MIN Total Thickness", x => x.TebalTotal_Min?.ToString("F2"), 6, "success"),
                    ("STD Total Thickness", x => x.TebalTotal_Asli?.ToString("F2"), 6, "primary"),
                    ("MAX Total Thickness", x => x.TebalTotal_Max?.ToString("F2"), 6, "danger"),
                    ("MIN Selisih", x => x.SelisihTebal_Min?.ToString("F2"), 6, "success"),
                    ("STD Selisih", x => x.SelisihTebal_Asli?.ToString("F2"), 6, "primary"),
                    ("MAX Selisih", x => x.SelisihTebal_Max?.ToString("F2"), 6, "danger"),

                    ("In Target", x => x.InnerTarget, 6, ""),
                    ("In Tol", x => x.InnerTol, 6, ""),
                    ("In LCL", x => x.InnerLCL, 6, ""),
                    ("In Min", x => x.InnerMin, 6, ""),
                    ("In UCL", x => x.InnerUCL, 6, ""),
                    ("In Max", x => x.InnerMax, 6, ""),

                    ("Mid Target", x => x.InnerMidTarget, 6, "info"),
                    ("Mid Tol", x => x.InnerMidTol, 6, "info"),
                    ("Mid LCL", x => x.InnerMidLCL, 6, "info"),
                    ("Mid Min", x => x.InnerMidMin, 6, "info"),
                    ("Mid UCL", x => x.InnerMidUCL, 6, "info"),
                    ("Mid Max", x => x.InnerMidMax, 6, "info"),

                    ("Th Target", x => x.ThickTarget, 6, ""),
                    ("Th Tol", x => x.ThickTol, 6, ""),
                    ("Th LCL", x => x.ThickLCL, 6, ""),
                    ("Th Min", x => x.ThickMin, 6, ""),
                    ("Th UCL", x => x.ThickUCL, 6, ""),
                    ("Th Max", x => x.ThickMax, 6, ""),

                    ("Tot Target", x => x.TotalTarget, 6, ""),
                    ("Tot Tol", x => x.TotalTol, 6, ""),
                    ("Tot LCL", x => x.TotalLCL, 6, ""),
                    ("Tot Min", x => x.TotalMin, 6, ""),
                    ("Tot UCL", x => x.TotalUCL, 6, ""),
                    ("Tot Max", x => x.TotalMax, 6, "")
                };

                var groups = new[] {
                    (Name: "Identifikasi Utama", Color: System.Drawing.Color.FromArgb(13, 110, 253)),
                    (Name: "Spesifikasi Hose & Dies", Color: System.Drawing.Color.FromArgb(230, 126, 34)),
                    (Name: "Mesh", Color: System.Drawing.Color.FromArgb(13, 110, 253)),
                    (Name: "Suhu Pengerjaan (°C)", Color: System.Drawing.Color.FromArgb(230, 126, 34)),
                    (Name: "Output Mesin", Color: System.Drawing.Color.FromArgb(13, 110, 253)),
                    (Name: "Proses Akhir", Color: System.Drawing.Color.FromArgb(230, 126, 34)),
                    (Name: "Final Quality Matrix", Color: System.Drawing.Color.FromArgb(13, 110, 253))
                };

                int currentCol = 1;
                for (int g = 0; g < groups.Length; g++)
                {
                    int count = colDefs.Count(c => c.G == g);
                    if (count > 0)
                    {
                        worksheet.Cells[1, currentCol, 1, currentCol + count - 1].Merge = true;
                        worksheet.Cells[1, currentCol].Value = groups[g].Name;
                        using (var r = worksheet.Cells[1, currentCol, 1, currentCol + count - 1])
                        {
                            r.Style.Font.Bold = true;
                            r.Style.Font.Color.SetColor(System.Drawing.Color.White);
                            r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            r.Style.Fill.BackgroundColor.SetColor(groups[g].Color);
                            r.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            r.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            r.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                        }
                        currentCol += count;
                    }
                }

                for (int i = 0; i < colDefs.Count; i++)
                {
                    var cell = worksheet.Cells[2, i + 1];
                    cell.Value = colDefs[i].H;
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    cell.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                    if (colDefs[i].Color == "success") {
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(210, 244, 234));
                        cell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(15, 81, 50));
                    } else if (colDefs[i].Color == "primary") {
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(207, 226, 255));
                        cell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(8, 66, 152));
                    } else if (colDefs[i].Color == "danger") {
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(248, 215, 218));
                        cell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(132, 32, 41));
                    } else if (colDefs[i].Color == "info") {
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(203, 240, 248));
                        cell.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(5, 81, 96));
                    } else {
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(240, 240, 240));
                        cell.Style.Font.Color.SetColor(System.Drawing.Color.Black);
                    }
                }

                int row = 3;
                foreach (var item in listData)
                {
                    for (int i = 0; i < colDefs.Count; i++)
                    {
                        var val = colDefs[i].Value(item);
                        worksheet.Cells[row, i + 1].Value = val;
                    }
                    row++;
                }
                
                worksheet.Cells.AutoFitColumns();
                
                var stream = new MemoryStream(package.GetAsByteArray());
                var fileName = $"SPS_Master_Export_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting SPS Master");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        [HttpPost]
        public async Task<IActionResult> ImportExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return Json(new { success = false, message = "File tidak valid" });
            }

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                var fileApprovalKey = BuildImportApprovalKey(stream);

                if (_approvalService.IsRequesterRole())
                {
                    var allowed = await _approvalService.HasConsumableApprovalAsync(ApprovalActionType.SpsImportTemplate, fileApprovalKey);
                    if (!allowed)
                    {
                        await _approvalService.CreateOrReusePendingRequestAsync(
                            ApprovalActionType.SpsImportTemplate,
                            fileApprovalKey,
                            $"Request import SPS template file '{file.FileName}'",
                            Url.Action(nameof(Index), "SpsMaster") ?? "/SpsMaster");

                        return Json(new
                        {
                            success = false,
                            status = "approval_required",
                            message = "Import ditahan. Request approval sudah dikirim ke inbox SUPERADMIN. Silakan tunggu approval sebelum import ulang file yang sama."
                        });
                    }
                }

                stream.Position = 0;
                
                using var package = new ExcelPackage(stream);
                
                // ========== SMART SHEET SELECTION ==========
                ExcelWorksheet? worksheet = null;
                var log = new System.Text.StringBuilder();
                log.AppendLine($"Total sheets: {package.Workbook.Worksheets.Count}");
                
                // Priority 1: Find sheet with SPS/Parameter/Digitalisasi keyword
                string[] sheetKeywords = { "Digitalisasi", "Parameter Setting", "SPS", "Master" };
                foreach (var ws in package.Workbook.Worksheets)
                {
                    log.AppendLine($"Sheet: '{ws.Name}' ({ws.Dimension?.Rows ?? 0} rows)");
                    
                    foreach (var keyword in sheetKeywords)
                    {
                        if (ws.Name.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                        {
                            worksheet = ws;
                            log.AppendLine($"âœ“ Selected sheet: {ws.Name} (matched: {keyword})");
                            _logger.LogInformation(log.ToString());
                            break;
                        }
                    }
                    if (worksheet != null) break;
                }
                
                // Priority 2: Largest sheet
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets
                        .OrderByDescending(w => w.Dimension?.Rows ?? 0)
                        .FirstOrDefault();
                    log.AppendLine($"âœ“ Selected largest sheet: {worksheet?.Name}");
                    _logger.LogInformation(log.ToString());
                }
                
                if (worksheet == null)
                {
                    return Json(new { success = false, message = "Worksheet tidak ditemukan" });
                }

                int rowCount = worksheet.Dimension?.Rows ?? 0;
                int colCount = worksheet.Dimension?.Columns ?? 0;
                int importedCount = 0;
                int importedDocsCount = 0;
                int updatedDocsCount = 0;
                int errorCount = 0;
                int skipped = 0;
                var errors = new List<string>();

                _logger.LogInformation($"Import Excel from sheet '{worksheet.Name}': {rowCount} rows, {colCount} columns");

                // ========== DETECT HEADER AND DATA START ROW ==========
                int headerRow = 1;
                int headerRow2 = 0; // Secondary header row (for CHS 3 Layer / Legacy)
                int dataStartRow = 2;
                
                // Scan row 1 to 3 across columns to detect titles and header positions
                string GetRowScanText(int r, int maxCols = 60)
                {
                    var texts = new List<string>();
                    for (int c = 1; c <= Math.Min(colCount, maxCols); c++)
                    {
                        var t = worksheet.Cells[r, c].Text?.Trim();
                        if (!string.IsNullOrWhiteSpace(t)) texts.Add(t.ToUpper());
                    }
                    return string.Join(" ", texts);
                }

                var row1Text = GetRowScanText(1);
                var row2Text = GetRowScanText(2);
                var row3Text = GetRowScanText(3);
                
                _logger.LogInformation($"Row 1 Scanned Text: '{row1Text}'");
                _logger.LogInformation($"Row 2 Scanned Text: '{row2Text}'");
                _logger.LogInformation($"Row 3 Scanned Text: '{row3Text}'");
                
                // Format 1: NON-CHS Single/Double Layer (header di row 3)
                if (row1Text.Contains("NON-CHS") || row1Text.Contains("SINGLE") || row1Text.Contains("DOUBLE LAYER PARAMETER"))
                {
                    _logger.LogInformation("Detected NON-CHS Double Layer format");
                    headerRow = 3;  // Headers: ID, No, Machine, No. Document, ...
                    headerRow2 = 5; // Secondary headers (ITEM, dll) ada di row 5
                    dataStartRow = 6;
                }
                // Format 2: CHS 2 Layer / 3 Layer Digitalisasi (header di row 3 + row 6 untuk Item)
                else if (row1Text.Contains("CHS 2 LAYER") || row1Text.Contains("CHS 3 LAYER") || row1Text.Contains("DIGITALISASI"))
                {
                    _logger.LogInformation($"Detected CHS Digitalisasi format - Title: {row1Text}");
                    headerRow = 3;  // Main headers di row 3
                    headerRow2 = 6; // Secondary headers (Item List) di row 6 untuk CHS 3 Layer
                    dataStartRow = 7; // Row 6 = numbers, Row 7+ = data
                }
                // Format 3: Template SPS Full (group labels di row 1, headers di row 2)
                else if (row1Text.Contains("IDENTIFIKASI") || row1Text.Contains("DOKUMEN") || row1Text.Contains("MATERIAL") || row1Text.Contains("DIES") ||
                         row2Text.Contains("NO. DOC") || row2Text.Contains("NO.DOC") || row2Text.Contains("ID EXCEL") || (row2Text.Contains("CUSTOMER") && row2Text.Contains("MACHINE")))
                {
                    _logger.LogInformation("Detected Template SPS Full format (row 1 = group labels, row 2 = headers)");
                    headerRow = 2;  // Headers: ID, No, MC, NO. DOC, ...
                    dataStartRow = 3; // Row 3+ = data
                }
                // Format 4: 3-Row Header (e.g. general masterlist)
                else if (row3Text.Contains("NO. DOC") || row3Text.Contains("NO.DOC") || row3Text.Contains("ID EXCEL"))
                {
                    _logger.LogInformation("Detected 3-row header format");
                    headerRow = 3;
                    dataStartRow = 4;
                }
                // Format 5: Simple format (header di row 1)
                else
                {
                    _logger.LogInformation("Detected simple format (header in row 1)");
                    headerRow = 1;
                    dataStartRow = 2;
                }
                
                _logger.LogInformation($"Using headerRow={headerRow}, headerRow2={headerRow2}, dataStartRow={dataStartRow}");

                // ========== FORMAT DETECTION ==========
                bool isDig3L = row1Text.Contains("CHS 3 LAYER");
                bool isDig2L = row1Text.Contains("CHS 2 LAYER");
                bool isLegacy = row1Text.Contains("NON-CHS") || row1Text.Contains("SINGLE") || row1Text.Contains("DOUBLE LAYER PARAMETER");
                bool isTemplateFull = headerRow == 2 || row1Text.Contains("IDENTIFIKASI") || row1Text.Contains("DOKUMEN") || row1Text.Contains("MATERIAL") || row1Text.Contains("DIES");
                
                _logger.LogInformation($"Format Detection: isDig3L={isDig3L}, isDig2L={isDig2L}, isLegacy={isLegacy}, isTemplateFull={isTemplateFull}");
                
                // ========== HARDCODED COLUMN INDICES (EPPlus is 1-indexed, DataTable was 0-indexed) ==========
                // Based on proven working backup logic, adjusted for EPPlus (+1 from backup)
                int idxDimensi, idxItem, idxInner, idxOuter, idxMiddle = 0;
                int idxRev, idxRevDate, idxCustomer, idxFormulasi;
                int idxYarn, idxPitchYarn, idxTensionIn, idxTensionOut;
                int idxNipple, idxTubeDie, idxMiddleDie, idxCoverDie, idxSpacerDie, idxADistance;
                int idxMeshDim1, idxMeshScreen1, idxMeshDim2, idxMeshScreen2, idxMeshDim3 = 0, idxMeshScreen3 = 0;
                int idxUseLimitsInner, idxUseLimitsOuter, idxUseLimitsMiddle = 0;
                int idxHeadTemp1, idxCyl11, idxCyl21, idxCyl31, idxScrewTemp1;
                int idxHeadTemp2, idxCyl12, idxCyl22, idxCyl32, idxScrewTemp2;
                int idxHeadTemp3 = 0, idxCyl13 = 0, idxCyl23 = 0, idxCyl33 = 0, idxScrewTemp3 = 0;
                int idxScrewSpeed1, idxScrewSpeed2, idxScrewSpeed3 = 0;
                int idxFeedRoll1, idxFeedRoll2, idxFeedRoll3 = 0;
                int idxPressure1, idxPressure2, idxPressure3 = 0;
                int idxCurrentValue = 0, idxAmMeter = 0, idxAmMeter2 = 0, idxAmMeter3 = 0;
                int idxPresetValue = 0, idxControlValue = 0;
                int idxSpiralPitchSetting = 0, idxSpiralPitchDisplay = 0, idxSpiralSpeed = 0, idxHoseSpeed = 0;
                int idxUnsmoothSurface = 0, idxMarkingSort = 0, idxTextMarkingMaterial = 0, idxMarkingColour = 0;
                int idxChillerWaterTemp = 0, idxDancerPosition = 0, idxCaterpillarGap = 0;
                int idxTakeUpConveyorSpeed = 0, idxCoolConveyorSpeed = 0, idxCoolConveyorSpeed2 = 0, idxConveyorRatio = 0;
                int idxToleranceInner = 0, idxToleranceOuter = 0;
                int idxTebalInner = 0, idxTebalInnerMiddle = 0, idxTebalTotal = 0, idxSelisihTebal = 0;
                int idxThickTarget = 0, idxThickTol = 0, idxThickLCL = 0, idxThickMin = 0, idxThickUCL = 0, idxThickMax = 0;
                int idxInnerTarget = 0, idxInnerTol = 0, idxInnerLCL = 0, idxInnerMin = 0, idxInnerUCL = 0, idxInnerMax = 0;
                int idxInnerMidTarget = 0, idxInnerMidTol = 0, idxInnerMidLCL = 0, idxInnerMidMin = 0, idxInnerMidUCL = 0, idxInnerMidMax = 0;
                int idxTotalTarget = 0, idxTotalTol = 0, idxTotalLCL = 0, idxTotalMin = 0, idxTotalUCL = 0, idxTotalMax = 0;
                int idxFeed1 = 0, idxFeed2 = 0, idxFeed3 = 0, idxTebalOuter = 0, idxODSensor = 0, idxCuttingSpeed = 0;
                
                if (isDig3L)
                {
                    _logger.LogInformation("Using CHS 3 Layer HARDCODED indices");
                    idxDimensi = 12; // Backup col 11 â†’ EPPlus col 12
                    idxItem = 108; // Col 108
                    idxRev = 7; idxRevDate = 9; idxCustomer = 8; idxFormulasi = 10; // Backup: 6,8,7,9 â†’ EPPlus: 7,9,8,10
                    idxInner = 14; idxMiddle = 15; idxOuter = 16;
                    idxUseLimitsInner = 17; idxUseLimitsMiddle = 18; idxUseLimitsOuter = 19;
                    idxYarn = 20; idxTensionIn = 21; idxTensionOut = 22;
                    idxNipple = 23; idxTubeDie = 24; idxMiddleDie = 25; idxCoverDie = 26;
                    idxSpacerDie = 27; idxADistance = 28;
                    idxMeshDim1 = 29; idxMeshScreen1 = 30;
                    idxMeshDim2 = 31; idxMeshScreen2 = 32;
                    idxMeshDim3 = 33; idxMeshScreen3 = 34;
                    idxPitchYarn = 0; // CHS 3 Layer doesn't have PitchYarn
                    // Temperature/Cylinder parameters
                    idxHeadTemp1 = 35; idxCyl11 = 36; idxCyl21 = 37; idxCyl31 = 38; idxScrewTemp1 = 39;
                    idxHeadTemp2 = 40; idxCyl12 = 41; idxCyl22 = 42; idxCyl32 = 43; idxScrewTemp2 = 44;
                    idxHeadTemp3 = 45; idxCyl13 = 46; idxCyl23 = 47; idxCyl33 = 48; idxScrewTemp3 = 49;
                    idxScrewSpeed1 = 50; idxScrewSpeed2 = 51; idxScrewSpeed3 = 52;
                    idxFeedRoll1 = 53; idxFeedRoll2 = 54; idxFeedRoll3 = 55;
                    idxPressure1 = 56; idxPressure2 = 57; idxPressure3 = 58;
                    // Machine parameters (Backup: 58-67 â†’ EPPlus: 59-68)
                    idxAmMeter = 59; idxAmMeter2 = 60; idxAmMeter3 = 61;
                    idxPresetValue = 62; idxControlValue = 63;
                    idxSpiralPitchSetting = 64; idxSpiralPitchDisplay = 65;
                    idxSpiralSpeed = 66; idxHoseSpeed = 67; idxUnsmoothSurface = 68;
                    // Marking/Conveyor parameters (Backup: 68-76 â†’ EPPlus: 69-77)
                    idxMarkingSort = 69; idxTextMarkingMaterial = 70; idxMarkingColour = 71;
                    idxChillerWaterTemp = 72; idxCaterpillarGap = 73;
                    idxTakeUpConveyorSpeed = 74; idxCoolConveyorSpeed = 75; idxCoolConveyorSpeed2 = 76;
                    idxConveyorRatio = 77;
                    // Tolerance/Thickness (Backup: 77-82 â†’ EPPlus: 78-83)
                    idxToleranceInner = 78; idxToleranceOuter = 79;
                    idxTebalInner = 80; idxTebalInnerMiddle = 81; idxTebalTotal = 82; idxSelisihTebal = 83;
                    // Quality Matrix (FIXED: Col 84-107)
                    // Col 84 (CF): Dimensi Inner Target, Col 85-89: Tol/LCL/Min/UCL/Max
                    idxInnerTarget = 84; idxInnerTol = 85; idxInnerLCL = 86;
                    idxInnerMin = 87; idxInnerUCL = 88; idxInnerMax = 89;
                    // Col 90 (CL): Inner Thick Target, Col 91-95: Tol/LCL/Min/UCL/Max
                    idxThickTarget = 90; idxThickTol = 91; idxThickLCL = 92;
                    idxThickMin = 93; idxThickUCL = 94; idxThickMax = 95;
                    // Col 96 (CR): Inner+Middle Target, Col 97-101: Tol/LCL/Min/UCL/Max
                    idxInnerMidTarget = 96; idxInnerMidTol = 97; idxInnerMidLCL = 98;
                    idxInnerMidMin = 99; idxInnerMidUCL = 100; idxInnerMidMax = 101;
                    // Col 102 (CX): Total Thick Target, Col 103-107: Tol/LCL/Min/UCL/Max
                    idxTotalTarget = 102; idxTotalTol = 103; idxTotalLCL = 104;
                    idxTotalMin = 105; idxTotalUCL = 106; idxTotalMax = 107;
                }
                else if (isDig2L)
                {
                    _logger.LogInformation("Using CHS 2 Layer HARDCODED indices");
                    idxDimensi = 12; // Backup col 11 â†’ EPPlus col 12
                    idxItem = 90; // Col 90
                    idxRev = 7; idxRevDate = 9; idxCustomer = 8; idxFormulasi = 10; // Backup: 6,8,7,9 â†’ EPPlus: 7,9,8,10
                    idxInner = 14; idxOuter = 15; // NO MIDDLE!
                    idxUseLimitsInner = 16; idxUseLimitsOuter = 17;
                    idxYarn = 18; idxPitchYarn = 19;
                    idxTensionIn = 20; idxTensionOut = 21;
                    idxNipple = 22; idxTubeDie = 23; idxMiddleDie = 24; idxCoverDie = 25;
                    idxSpacerDie = 26; idxADistance = 27;
                    idxMeshDim1 = 28; idxMeshScreen1 = 29;
                    idxMeshDim2 = 30; idxMeshScreen2 = 31;
                    idxMeshDim3 = 0; idxMeshScreen3 = 0; // CHS 2 Layer only has 2 mesh screens
                    // Temperature/Cylinder parameters
                    idxHeadTemp1 = 32; idxCyl11 = 33; idxCyl21 = 34; idxCyl31 = 35; idxScrewTemp1 = 36;
                    idxHeadTemp2 = 37; idxCyl12 = 38; idxCyl22 = 39; idxCyl32 = 40; idxScrewTemp2 = 41;
                    idxScrewSpeed1 = 42; idxScrewSpeed2 = 43;
                    idxFeedRoll1 = 44; idxFeedRoll2 = 45;
                    idxPressure1 = 46; idxPressure2 = 47;
                    // Machine parameters (Backup: 47-56 â†’ EPPlus: 48-57)
                    idxCurrentValue = 48; idxAmMeter = 49; idxAmMeter2 = 50;
                    idxPresetValue = 51; idxControlValue = 52;
                    idxSpiralPitchSetting = 53; idxSpiralPitchDisplay = 54;
                    idxSpiralSpeed = 55; idxHoseSpeed = 56; idxUnsmoothSurface = 57;
                    // Marking/Conveyor parameters (Backup: 57-65 â†’ EPPlus: 58-66)
                    idxMarkingSort = 58; idxTextMarkingMaterial = 59; idxMarkingColour = 60;
                    idxChillerWaterTemp = 61; idxDancerPosition = 62; idxCaterpillarGap = 63;
                    idxTakeUpConveyorSpeed = 64; idxCoolConveyorSpeed = 65; idxConveyorRatio = 66;
                    // Tolerance/Thickness (Backup: 66-70 â†’ EPPlus: 67-71)
                    idxToleranceInner = 67; idxToleranceOuter = 68;
                    idxTebalInner = 69; idxTebalTotal = 70; idxSelisihTebal = 71;
                    // Quality Matrix CHS 2L (FIXED: Col 72-89)
                    // Col 72 (BT): Inner Diameter Target, Col 73-77: Tol/LCL/Min/UCL/Max
                    idxInnerTarget = 72; idxInnerTol = 73; idxInnerLCL = 74;
                    idxInnerMin = 75; idxInnerUCL = 76; idxInnerMax = 77;
                    // Col 78 (BZ): Inner Thickness Target, Col 79-83: Tol/LCL/Min/UCL/Max
                    idxThickTarget = 78; idxThickTol = 79; idxThickLCL = 80;
                    idxThickMin = 81; idxThickUCL = 82; idxThickMax = 83;
                    // Col 84 (CF): Total Thickness Target, Col 85-89: Toleransi/LCL/Min/UCL/Max
                    idxTotalTarget = 84; idxTotalTol = 85; idxTotalLCL = 86;
                    idxTotalMin = 87; idxTotalUCL = 88; idxTotalMax = 89;
                }
                else if (isTemplateFull)
                {
                    _logger.LogInformation("Using Template SPS Full HARDCODED indices");
                    idxDimensi = 9;
                    idxItem = 302;
                    idxRev = 5; idxRevDate = 6; idxCustomer = 7; idxFormulasi = 10;
                    idxInner = 13; idxMiddle = 15; idxOuter = 14;
                    idxUseLimitsInner = 16; idxUseLimitsOuter = 17; idxUseLimitsMiddle = 18;
                    idxNipple = 19; idxTubeDie = 23; idxCoverDie = 27; idxMiddleDie = 31; idxSpacerDie = 35; idxADistance = 39;
                    idxYarn = 43; idxTensionIn = 44; idxTensionOut = 45;
                    idxMeshScreen1 = 46; idxMeshDim1 = 47;
                    idxMeshScreen2 = 51; idxMeshDim2 = 52;
                    idxMeshScreen3 = 56; idxMeshDim3 = 57;
                    idxPitchYarn = 61;
                    
                    // Temperature/Cylinder parameters
                    idxHeadTemp1 = 65; idxCyl11 = 69; idxCyl21 = 73; idxCyl31 = 77; idxScrewTemp1 = 81; idxScrewSpeed1 = 85; idxPressure1 = 89;
                    idxHeadTemp2 = 93; idxCyl12 = 97; idxCyl22 = 101; idxCyl32 = 105; idxScrewTemp2 = 109; idxScrewSpeed2 = 113; idxPressure2 = 117;
                    idxHeadTemp3 = 121; idxCyl13 = 125; idxCyl23 = 129; idxCyl33 = 133; idxScrewTemp3 = 137; idxScrewSpeed3 = 141; idxPressure3 = 145;
                    
                    // Machine parameters
                    idxCurrentValue = 149;
                    idxFeed1 = 150; idxFeed2 = 154; idxFeed3 = 158;
                    idxFeedRoll1 = 162; idxFeedRoll2 = 166; idxFeedRoll3 = 170;
                    idxAmMeter = 174; idxAmMeter2 = 178; idxAmMeter3 = 182;
                    idxPresetValue = 186; idxControlValue = 190;
                    idxSpiralPitchSetting = 194; idxSpiralPitchDisplay = 198;
                    idxSpiralSpeed = 202; idxHoseSpeed = 206;
                    idxChillerWaterTemp = 210; idxDancerPosition = 214; idxCaterpillarGap = 218;
                    idxTakeUpConveyorSpeed = 226; idxCoolConveyorSpeed = 230; idxCoolConveyorSpeed2 = 234; idxConveyorRatio = 238;
                    idxODSensor = 242; idxCuttingSpeed = 222;
                    
                    // Tolerance/Thickness
                    idxToleranceInner = 250; idxToleranceOuter = 254;
                    idxTebalInner = 258; idxTebalOuter = 262; idxTebalInnerMiddle = 266; idxTebalTotal = 270; idxSelisihTebal = 274;
                    
                    // Quality Matrix
                    idxInnerTarget = 278; idxInnerTol = 279; idxInnerLCL = 280;
                    idxInnerMin = 281; idxInnerUCL = 282; idxInnerMax = 283;
                    
                    idxInnerMidTarget = 284; idxInnerMidTol = 285; idxInnerMidLCL = 286;
                    idxInnerMidMin = 287; idxInnerMidUCL = 288; idxInnerMidMax = 289;
                    
                    idxThickTarget = 290; idxThickTol = 291; idxThickLCL = 292;
                    idxThickMin = 293; idxThickUCL = 294; idxThickMax = 295;
                    
                    idxTotalTarget = 296; idxTotalTol = 297; idxTotalLCL = 298;
                    idxTotalMin = 299; idxTotalUCL = 300; idxTotalMax = 301;
                }
                else if (isLegacy || (headerRow == 3 && worksheet.Cells[3, 43].Text?.Trim().Contains("Â± Inner") == true))
                {
                    // Double Layer_Digitalisasi or NON-CHS format
                    _logger.LogWarning("ðŸ” FORMAT DETECTED: Double Layer_Digitalisasi/Non-CHS");
                    _logger.LogWarning($"   headerRow={headerRow}, Cell[3,43]='{worksheet.Cells[3, 43].Text}'");
                    idxDimensi = 10; // Col 10: Dimensi
                    idxItem = 51; // Assuming item list around col 51-52
                    idxRev = 5; idxRevDate = 7; idxCustomer = 6; idxFormulasi = 8; // Row 3 headers
                    idxInner = 12; idxOuter = 13; idxMiddle = 0; // No middle tube in Double Layer
                    idxUseLimitsInner = 14; idxUseLimitsOuter = 15; idxUseLimitsMiddle = 0;
                    idxNipple = 16; idxTubeDie = 17; idxCoverDie = 18;
                    idxMeshScreen1 = 19; idxMeshScreen2 = 20;
                    idxYarn = 0; idxPitchYarn = 0; idxTensionIn = 0; idxTensionOut = 0;
                    idxMiddleDie = 0; idxSpacerDie = 0; idxADistance = 0;
                    idxMeshDim1 = 0; idxMeshDim2 = 0; idxMeshDim3 = 0; idxMeshScreen3 = 0;
                    // Temperature indices
                    idxHeadTemp1 = 21; idxHeadTemp2 = 22;
                    idxCyl11 = 23; idxCyl21 = 25; idxCyl31 = 0; 
                    idxCyl12 = 24; idxCyl22 = 26; idxCyl32 = 0;
                    idxScrewTemp1 = 29; idxScrewTemp2 = 30;
                    idxScrewSpeed1 = 31; idxScrewSpeed2 = 32;
                    idxFeedRoll1 = 27; idxFeedRoll2 = 28;
                    idxPressure1 = 33; idxPressure2 = 34;
                    idxAmMeter = 35; idxCurrentValue = 0; idxAmMeter2 = 0;
                    idxPresetValue = 0; idxControlValue = 0;
                    idxSpiralPitchSetting = 0; idxSpiralPitchDisplay = 0;
                    idxSpiralSpeed = 0; idxHoseSpeed = 0; idxUnsmoothSurface = 0;
                    idxMarkingSort = 37; idxTextMarkingMaterial = 38; idxMarkingColour = 39;
                    idxChillerWaterTemp = 40; idxDancerPosition = 0; idxCaterpillarGap = 0;
                    idxTakeUpConveyorSpeed = 42; idxCoolConveyorSpeed = 0; idxCoolConveyorSpeed2 = 0; idxConveyorRatio = 0;
                    // TOLERANCE/THICKNESS FOR DOUBLE LAYER DIGITALISASI (Col 43-48)
                    idxToleranceInner = 43; // Â± Inner
                    idxToleranceOuter = 44;  // Â± Outer
                    idxTebalInner = 45;      // Tebal Inner
                    idxTebalTotal = 47;      // Tebal Total
                    idxSelisihTebal = 48;    // Selisih tebal
                    idxTebalInnerMiddle = 0; // No middle in Double Layer
                    _logger.LogWarning($"   ðŸŽ¯ TOLERANCE COLUMNS: Inner={idxToleranceInner}, Outer={idxToleranceOuter}, TebalInner={idxTebalInner}, TebalTotal={idxTebalTotal}");
                }
                else
                {
                    _logger.LogInformation("Using Legacy/Simple format indices");
                    idxDimensi = 10; // Backup col 9 â†’ EPPlus col 10
                    idxItem = 51; // Assuming legacy format
                    idxRev = 5; idxRevDate = 7; idxCustomer = 6; idxFormulasi = 8; // Backup: 4,6,5,7 â†’ EPPlus: 5,7,6,8
                    idxInner = 12; idxOuter = 13;
                    idxUseLimitsInner = 14; idxUseLimitsOuter = 15;
                    idxNipple = 16; idxTubeDie = 17; idxCoverDie = 18;
                    idxMeshScreen1 = 19; idxMeshScreen2 = 20;
                    // Legacy format has fewer columns
                    idxYarn = 0; idxPitchYarn = 0; idxTensionIn = 0; idxTensionOut = 0;
                    idxMiddleDie = 0; idxSpacerDie = 0; idxADistance = 28;
                    idxMeshDim1 = 0; idxMeshDim2 = 0; idxMeshDim3 = 0; idxMeshScreen3 = 0;
                    // Legacy temperature indices
                    idxHeadTemp1 = 21; idxHeadTemp2 = 22;
                    idxCyl11 = 0; idxCyl21 = 0; idxCyl31 = 0; idxScrewTemp1 = 0;
                    idxCyl12 = 0; idxCyl22 = 0; idxCyl32 = 0; idxScrewTemp2 = 0;
                    idxScrewSpeed1 = 0; idxScrewSpeed2 = 0;
                    idxFeedRoll1 = 0; idxFeedRoll2 = 0;
                    idxPressure1 = 0; idxPressure2 = 0;
                }
                
                _logger.LogInformation($"Column Indices: Item={idxItem}, Inner={idxInner}, Outer={idxOuter}, Middle={idxMiddle}");
                _logger.LogInformation($"  Yarn={idxYarn}, PitchYarn={idxPitchYarn}, TensionIn={idxTensionIn}, TensionOut={idxTensionOut}");
                _logger.LogInformation($"  Nipple={idxNipple}, TubeDie={idxTubeDie}, MiddleDie={idxMiddleDie}, CoverDie={idxCoverDie}");

                // ========== SMART COLUMN DETECTION (Legacy compatibility) ==========
                var headers = new Dictionary<string, int>();
                var headerDebugLog = new System.Text.StringBuilder();
                headerDebugLog.AppendLine($"Reading headers from row {headerRow}:");
                
                // SCAN PRIMARY HEADER (row 3 for CHS, or row 1 for simple)
                for (int col = 1; col <= colCount; col++)
                {
                    var headerValue = worksheet.Cells[headerRow, col].Text?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(headerValue) && !headerValue.All(char.IsDigit)) // Skip column number row
                    {
                        // Store with normalized key
                        var normalizedKey = headerValue.ToUpper()
                            .Replace("|", "")
                            .Replace("#", "")
                            .Replace(".", "")
                            .Replace(" ", "")
                            .Replace("_", "")
                            .Replace("-", "");
                        headers[normalizedKey] = col;
                        headers[headerValue.ToUpper()] = col;
                        
                        // Log first 20 columns for debugging
                        if (col <= 20)
                        {
                            headerDebugLog.AppendLine($"  Col {col}: '{headerValue}' â†’ normalized: '{normalizedKey}'");
                        }
                    }
                }
                
                // SCAN SECONDARY HEADER (row 6 for CHS 3 Layer - untuk kolom "Item")
                if (headerRow2 > 0)
                {
                    headerDebugLog.AppendLine($"\nReading secondary headers from row {headerRow2}:");
                    for (int col = 1; col <= colCount; col++)
                    {
                        var headerValue = worksheet.Cells[headerRow2, col].Text?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(headerValue) && !headerValue.All(char.IsDigit))
                        {
                            var normalizedKey = headerValue.ToUpper()
                                .Replace("|", "")
                                .Replace("#", "")
                                .Replace(".", "")
                                .Replace(" ", "")
                                .Replace("_", "")
                                .Replace("-", "");
                            
                            // Don't overwrite existing headers, only add new ones
                            if (!headers.ContainsKey(normalizedKey))
                            {
                                headers[normalizedKey] = col;
                                headers[headerValue.ToUpper()] = col;
                                
                                headerDebugLog.AppendLine($"  Col {col}: '{headerValue}' â†’ normalized: '{normalizedKey}'");
                            }
                        }
                    }
                }
                
                // SPECIAL CHECK: Always check last 20 columns for "Item" keyword (CHS 2/3 Layer has Item at end)
                headerDebugLog.AppendLine($"\nChecking last 20 columns for 'ITEM' keyword:");
                for (int col = Math.Max(1, colCount - 19); col <= colCount; col++)
                {
                    // Check both header rows
                    var val3 = worksheet.Cells[headerRow, col].Text?.Trim().ToUpper() ?? "";
                    var val6 = headerRow2 > 0 ? worksheet.Cells[headerRow2, col].Text?.Trim().ToUpper() ?? "" : "";
                    
                    if (val3.Contains("ITEM") || val6.Contains("ITEM"))
                    {
                        var headerValue = !string.IsNullOrEmpty(val3) && val3.Contains("ITEM") ? val3 : val6;
                        headers["ITEM"] = col;
                        headers["ITEMLIST"] = col;
                        headerDebugLog.AppendLine($"  âœ“ Col {col}: ITEM FOUND! (value: '{headerValue}')");
                    }
                }
                
                _logger.LogInformation($"Found {headers.Count} header columns");
                _logger.LogInformation(headerDebugLog.ToString());

                // Helper function: Find LEFTMOST column with multiple possible names
                // For CHS format, NO. DOC has 2-3 columns, we want the leftmost (smallest column number)
                int FindColumn(params string[] possibleNames)
                {
                    _logger.LogInformation($"Finding column for: {string.Join(", ", possibleNames)}");
                    
                    var matchedColumns = new List<int>();
                    
                    // Collect all exact matches first
                    foreach (var name in possibleNames)
                    {
                        var normalized = name.ToUpper()
                            .Replace("|", "")
                            .Replace("#", "")
                            .Replace(".", "")
                            .Replace(" ", "")
                            .Replace("_", "")
                            .Replace("-", "");
                        
                        // Try exact normalized match
                        if (headers.ContainsKey(normalized))
                        {
                            matchedColumns.Add(headers[normalized]);
                            _logger.LogInformation($"  â†’ Found at col {headers[normalized]} (normalized match: '{normalized}')");
                        }
                        
                        // Try exact uppercase match
                        if (headers.ContainsKey(name.ToUpper()))
                        {
                            matchedColumns.Add(headers[name.ToUpper()]);
                            _logger.LogInformation($"  â†’ Found at col {headers[name.ToUpper()]} (exact match: '{name.ToUpper()}')");
                        }
                    }
                    
                    // Try partial match if no exact match found
                    if (matchedColumns.Count == 0)
                    {
                        foreach (var name in possibleNames)
                        {
                            var normalized = name.ToUpper()
                                .Replace("|", "")
                                .Replace("#", "")
                                .Replace(".", "")
                                .Replace(" ", "")
                                .Replace("_", "")
                                .Replace("-", "");
                            
                            foreach (var kvp in headers)
                            {
                                var headerKey = kvp.Key;
                                
                                // Check if either contains the other
                                if ((headerKey.Length >= 3 && normalized.Length >= 3) &&
                                    (headerKey.Contains(normalized) || normalized.Contains(headerKey)))
                                {
                                    matchedColumns.Add(kvp.Value);
                                    _logger.LogInformation($"  â†’ Found at col {kvp.Value} (partial match: '{headerKey}')");
                                }
                            }
                        }
                    }
                    
                    // Return LEFTMOST column (smallest column number)
                    if (matchedColumns.Count > 0)
                    {
                        var leftmost = matchedColumns.Min();
                        if (matchedColumns.Count > 1)
                        {
                            _logger.LogInformation($"  âœ“ Multiple columns found ({string.Join(", ", matchedColumns.OrderBy(c => c))}), using LEFTMOST: col {leftmost}");
                        }
                        return leftmost;
                    }
                    
                    _logger.LogWarning($"  â†’ Column NOT FOUND for: {string.Join(", ", possibleNames)}");
                    return 0;
                }

                // Helper function: Get cell value safely
                string GetCellValue(int row, int col)
                {
                    if (col <= 0) return "";
                    // Use .Value instead of .Text to preserve special characters like Â±
                    var cellValue = worksheet.Cells[row, col].Value;
                    if (cellValue == null) return "";
                    
                    var value = cellValue.ToString()?.Trim() ?? "";
                    // Debug: log unicode codepoints if value contains potential ± character
                    if (value.Length > 0 && value.Any(c => c == '\u00B1' || c == '\u00C2' || c > 127))
                    {
                        var codepoints = string.Join("-", value.Select(c => $"U+{(int)c:X4}"));
                        _logger.LogDebug($"GetCellValue[{row},{col}] special chars: [{value}] codepoints={codepoints}");
                    }
                    return (string.IsNullOrWhiteSpace(value) || value == "-") ? "" : value;
                }

                // Find critical columns (basic metadata)
                int idExcelCol = FindColumn("ID", "ID EXCEL", "EXCEL ID", "EXCELID", "NO");
                int docNumberCol = FindColumn("NO. DOC", "NO.DOC", "NODOC", "DOCUMENT NUMBER", "NO. DOCUMENT", "DOCUMENTNUMBER", "NO DOCUMENT", "DOC NUMBER", "DOC NO", "DOCNO", "SOP NUMBER", "PROD NUMBER", "VI-SOP-PROD", "VISOPPROD", "DOC", "DOCUMENT");
                int machineCol = FindColumn("MACHINE", "MESIN");
                int machineCodeCol = FindColumn("MC", "MACHINE CODE", "MACHINECODE", "KODE MESIN", "M/C");
                
                // ITEM LIST - Try to find dynamically first, fallback to HARDCODED
                int itemListCol = FindColumn("ITEM LIST", "ITEMLIST", "ITEM");
                if (itemListCol <= 0)
                {
                    itemListCol = idxItem;
                    _logger.LogInformation($"  â†’ Item column not found dynamically, using HARDCODED: col {itemListCol} (based on format detection)");
                }
                else
                {
                    _logger.LogInformation($"  â†’ Using DYNAMICALLY found Item column: col {itemListCol}");
                }
                
                int noCol = FindColumn("NO", "NUMBER", "NO.");
                int revNumberCol = FindColumn("NO. REV.", "NO.REV", "NOREV", "REV. NO", "REVISION NUMBER", "NO REV");
                int revDateCol = FindColumn("REV. DATE", "REV.DATE", "REVDATE", "REVISION DATE", "REVISIONDATE", "REVISION DATE");
                int customerCol = FindColumn("CUSTOMER", "CUST");
                int formulasiCol = FindColumn("FORMULASI", "FORMULA");
                int hoseTypeCol = FindColumn("HOSE TYPE", "HOSETYPE", "TYPE");
                int dimensiCol = FindColumn("DIMENSI", "DIMENSION", "DIM");
                int materialCol = FindColumn("MATERIAL TSM", "MATERIAL", "MATERIALTSM", "TSM");
                
                // NOTE: Inner/Outer/Middle Tube columns are now using HARDCODED indices (idxInner, idxOuter, idxMiddle)
                // Tolerance/Parameter columns (optional detection for non-standard formats)
                
                
                // NOTE: Inner/Outer/Middle Tube columns are now using HARDCODED indices (idxInner, idxOuter, idxMiddle)
                // Tolerance/Parameter columns (optional detection for non-standard formats)
                int toleranceInnerCol = FindColumn("TOLERANCE | #INNER", "TOLERANCE INNER", "TOL INNER", "TOLERANCEINNER");
                int toleranceOuterCol = FindColumn("TOLERANCE | #OUTER", "TOLERANCE OUTER", "TOL OUTER", "TOLERANCEOUTER");
                int toleranceMiddleCol = FindColumn("TOLERANCE | #MIDDLE", "TOLERANCE MIDDLE", "TOL MIDDLE", "TOLERANCEMIDDLE");
                int compoundCol = FindColumn("COMPOUND", "COMPOUND | #MATERIAL");
                int screenCol = FindColumn("SCREEN", "SCREEN | #MATERIAL");

                _logger.LogInformation($"Column Detection Results:");
                _logger.LogInformation($"  ID Excel: col {idExcelCol}");
                _logger.LogInformation($"  NO: col {noCol}");
                _logger.LogInformation($"  DocumentNumber: col {docNumberCol}");
                _logger.LogInformation($"  Machine: col {machineCol}");
                _logger.LogInformation($"  MachineCode: col {machineCodeCol}");
                _logger.LogInformation($"  ItemList: col {itemListCol} (HARDCODED)");
                _logger.LogInformation($"  Dimensi: col {idxDimensi} (HARDCODED)");
                _logger.LogInformation($"  InnerTube: col {idxInner} (HARDCODED)");
                _logger.LogInformation($"  MiddleTube: col {idxMiddle} (HARDCODED)");
                _logger.LogInformation($"  OuterCover: col {idxOuter} (HARDCODED)");
                _logger.LogInformation($"  Yarn: col {idxYarn} (HARDCODED)");
                _logger.LogInformation($"  PitchYarn: col {idxPitchYarn} (HARDCODED)");
                _logger.LogInformation($"  Nipple: col {idxNipple} (HARDCODED)");
                _logger.LogInformation($"  AmMeter: col {idxAmMeter}, AmMeter2: col {idxAmMeter2}, AmMeter3: col {idxAmMeter3} (HARDCODED)");
                _logger.LogInformation($"  PresetValue: col {idxPresetValue}, ControlValue: col {idxControlValue} (HARDCODED)");
                _logger.LogInformation($"  SpiralPitchSetting: col {idxSpiralPitchSetting}, SpiralSpeed: col {idxSpiralSpeed} (HARDCODED)");
                _logger.LogInformation($"  HoseSpeed: col {idxHoseSpeed}, UnsmoothSurface: col {idxUnsmoothSurface} (HARDCODED)");
                _logger.LogInformation($"  FeedRoll1: col {idxFeedRoll1}, FeedRoll2: col {idxFeedRoll2}, FeedRoll3: col {idxFeedRoll3} (HARDCODED)");
                _logger.LogInformation($"  Quality Matrix: Thick col {idxThickTarget}-{idxThickMax}, Inner col {idxInnerTarget}-{idxInnerMax}, InnerMid col {idxInnerMidTarget}-{idxInnerMidMax}, Total col {idxTotalTarget}-{idxTotalMax} (HARDCODED)");

                if (docNumberCol == 0)
                {
                    _logger.LogWarning("âŒ Kolom NO. DOC tidak ditemukan, akan gunakan ID EXCEL atau row number sebagai fallback");
                    
                    // Fallback: gunakan ID EXCEL column sebagai document number
                    if (idExcelCol > 0)
                    {
                        _logger.LogInformation("  â†’ Menggunakan ID EXCEL sebagai document number");
                        docNumberCol = idExcelCol;
                    }
                    else
                    {
                        // Last resort: gunakan row number
                        _logger.LogWarning("  â†’ ID EXCEL juga tidak ditemukan, akan gunakan row number");
                    }
                }
                
                // Jika masih tidak ada docNumberCol dan idExcelCol, show error
                if (docNumberCol == 0 && idExcelCol == 0)
                {
                    // Build helpful error message with what we found
                    var errorMsg = new System.Text.StringBuilder();
                    errorMsg.AppendLine("âŒ Kolom 'NO. DOC' dan 'ID EXCEL' tidak ditemukan!");
                    errorMsg.AppendLine($"\nðŸ“Š Info file:");
                    errorMsg.AppendLine($"- Sheet: {worksheet.Name}");
                    errorMsg.AppendLine($"- Total kolom: {colCount}");
                    errorMsg.AppendLine($"- Header row: {headerRow}");
                    errorMsg.AppendLine($"- Data start row: {dataStartRow}");
                    errorMsg.AppendLine($"\nðŸ“ Kolom yang berhasil dibaca (30 pertama):");
                    
                    int shown = 0;
                    foreach (var kvp in headers.OrderBy(h => h.Value).Take(30))
                    {
                        errorMsg.AppendLine($"  Col {kvp.Value}: '{kvp.Key}'");
                        shown++;
                    }
                    
                    if (headers.Count == 0)
                    {
                        errorMsg.AppendLine("  âš  TIDAK ADA HEADER YANG TERBACA!");
                        errorMsg.AppendLine($"\nðŸ’¡ Coba cek:");
                        errorMsg.AppendLine($"  - Apakah sheet '{worksheet.Name}' yang benar?");
                        errorMsg.AppendLine($"  - Apakah row {headerRow} berisi nama kolom?");
                    }
                    
                    _logger.LogError(errorMsg.ToString());
                    return Json(new { 
                        success = false, 
                        message = errorMsg.ToString().Replace("\n", "<br/>")
                    });
                }

                // ========== GROUP BY DOCUMENT NUMBER ==========
                var docGroups = new Dictionary<string, List<int>>();
                bool useRowNumberFallback = (docNumberCol == 0 && idExcelCol == 0);
                int consecutiveEmpty = 0;
                const int MAX_CONSECUTIVE_EMPTY = 100;
                
                _logger.LogInformation($"Starting data import from row {dataStartRow}");
                
                for (int row = dataStartRow; row <= rowCount; row++)
                {
                    string docNum;
                    
                    if (useRowNumberFallback)
                    {
                        // Fallback: gunakan row number sebagai doc number
                        docNum = $"ROW-{row}";
                        _logger.LogInformation($"Row {row}: Using row number as doc number: {docNum}");
                    }
                    else
                    {
                        docNum = GetCellValue(row, docNumberCol);
                    }
                    
                    // Skip empty, header, invalid, or Excel error rows
                    if (string.IsNullOrWhiteSpace(docNum) || 
                        docNum.ToUpper().Contains("DOC") || 
                        docNum.ToUpper().Contains("DOCUMENT") ||
                        docNum.StartsWith("#REF") ||
                        docNum.StartsWith("#N/A") ||
                        docNum.StartsWith("#VALUE") ||
                        docNum.StartsWith("#ERROR") ||
                        docNum.StartsWith("#") ||
                        docNum.StartsWith("ROW-") ||
                        docNum == "-" ||
                        docNum.Length < 3)
                    {
                        skipped++;
                        consecutiveEmpty++;
                        
                        // Log only every 1000 skipped rows
                        if (skipped % 1000 == 0)
                        {
                            _logger.LogInformation($"Skipped {skipped} invalid rows (row {row})");
                        }
                        
                        // Early termination
                        if (consecutiveEmpty >= MAX_CONSECUTIVE_EMPTY)
                        {
                            _logger.LogInformation($"Stopping: {consecutiveEmpty} consecutive empty rows at row {row}");
                            break;
                        }
                        continue;
                    }
                    
                    consecutiveEmpty = 0;

                    if (!docGroups.ContainsKey(docNum))
                    {
                        docGroups[docNum] = new List<int>();
                    }
                    docGroups[docNum].Add(row);
                }

                _logger.LogInformation($"Found {docGroups.Count} unique documents");

                // ========== PROCESS EACH ROW (1 RECORD = 1 ITEMLIST) ==========
                var allRows = docGroups.SelectMany(g => g.Value).ToList();
                foreach (var rowNum in allRows)
                {
                    try
                    {
                        int firstRow = rowNum; // Alias to keep existing code working
                        
                        // ====== LOCAL HELPER: Supports both ± format AND manual MIN/STD/MAX columns ======
                        // Untuk isTemplateFull/manual: setiap parameter punya 4 kolom: raw(N), MIN(N+1), STD(N+2), MAX(N+3)
                        // Jika kolom raw berisi format ± / asymmetric → parse otomatis
                        // Jika kolom raw kosong/tidak pakai ± → baca MIN/STD/MAX dari kolom terpisah (manual input)
                        bool TryParseCleanDecimal(string? str, out decimal val)
                        {
                            val = 0;
                            if (string.IsNullOrWhiteSpace(str)) return false;
                            var cleanedStr = str.Trim().Replace(" ", "").Replace(",", ".");
                            return decimal.TryParse(cleanedStr, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out val);
                        }

                        void AssignFullOrManual(int rawCol,
                            Action<decimal?> setMin, Action<decimal?> setAsli, Action<decimal?> setMax)
                        {
                            if (rawCol <= 0) { setMin(null); setAsli(null); setMax(null); return; }
                            var rawVal = GetCellValue(firstRow, rawCol);
                            var parsed = ParsePlusMinusValue(rawVal);
                            if (parsed.min != null || parsed.asli != null || parsed.max != null)
                            {
                                // Format ± berhasil di-parse → gunakan hasil parse
                                setMin(parsed.min); setAsli(parsed.asli); setMax(parsed.max);
                                return;
                            }

                            // Cek kolom manual MIN/STD/MAX (rawCol+1, rawCol+2, rawCol+3)
                            var minStr = GetCellValue(firstRow, rawCol + 1);
                            var stdStr = GetCellValue(firstRow, rawCol + 2);
                            var maxStr = GetCellValue(firstRow, rawCol + 3);

                            // Jika stdStr berisi format formula/toleransi (contoh: 7.3±0.1 atau 10.5+0.1/-0.2)
                            if (!string.IsNullOrWhiteSpace(stdStr))
                            {
                                var stdParsed = ParsePlusMinusValue(stdStr);
                                if (stdParsed.min != null || stdParsed.asli != null || stdParsed.max != null)
                                {
                                    decimal? minVal = stdParsed.min;
                                    decimal? maxVal = stdParsed.max;
                                    if (TryParseCleanDecimal(minStr, out var expMin)) minVal = expMin;
                                    if (TryParseCleanDecimal(maxStr, out var expMax)) maxVal = expMax;

                                    setMin(minVal); setAsli(stdParsed.asli); setMax(maxVal);
                                    return;
                                }
                            }

                            decimal? minV = TryParseCleanDecimal(minStr, out var mv) ? mv : (decimal?)null;
                            decimal? stdV = TryParseCleanDecimal(stdStr, out var sv) ? sv : (decimal?)null;
                            decimal? maxV = TryParseCleanDecimal(maxStr, out var xv) ? xv : (decimal?)null;

                            if (minV != null || stdV != null || maxV != null)
                            {
                                _logger.LogInformation($"  📋 Manual MIN/STD/MAX col {rawCol}: MIN={minV}, STD={stdV}, MAX={maxV}");
                                setMin(minV); setAsli(stdV); setMax(maxV);
                            }
                            else
                            {
                                setMin(null); setAsli(null); setMax(null);
                            }
                        }
                        string docNumber = useRowNumberFallback ? $"ROW-{firstRow}" : GetCellValue(firstRow, docNumberCol);
                        var itemListValue = itemListCol > 0 ? GetCellValue(firstRow, itemListCol) : "";

                        _logger.LogInformation($"Processing row {firstRow}: Doc {docNumber}");

                        // Check if document exists
                        var existingDoc = await _context.SpsNoDocs
                            .Include(s => s.ItemLists)
                            .FirstOrDefaultAsync(s => s.DocumentNumber == docNumber);

                        SpsNoDoc spsNoDoc;
                        bool isNewDoc = false;
                        
                        if (existingDoc != null)
                        {
                            // UPDATE existing document
                            _logger.LogInformation($"📝 UPDATING EXISTING DOC: '{docNumber}'");
                            spsNoDoc = existingDoc;
                        }
                        else
                        {
                            // CREATE new document
                            _logger.LogInformation($"✨ CREATING NEW DOC: '{docNumber}'");
                            spsNoDoc = new SpsNoDoc();
                            isNewDoc = true;
                        }
                            
                            // Update/Set all fields from Excel
                            {
                            spsNoDoc.DocumentNumber = docNumber;
                            spsNoDoc.No = GetCellValue(firstRow, noCol);
                            spsNoDoc.Machine = GetCellValue(firstRow, machineCol);
                            // Fallback: jika kolom MC/Machine Code tidak ditemukan, gunakan nilai Machine
                            var machineCodeVal = machineCodeCol > 0 ? GetCellValue(firstRow, machineCodeCol) : "";
                            spsNoDoc.MachineCode = !string.IsNullOrEmpty(machineCodeVal) ? machineCodeVal : GetCellValue(firstRow, machineCol);
                            spsNoDoc.RevisionNumber = GetCellValue(firstRow, idxRev); // HARDCODED
                            spsNoDoc.RevisionDate = GetCellValue(firstRow, idxRevDate); // HARDCODED
                            spsNoDoc.Customer = GetCellValue(firstRow, idxCustomer); // HARDCODED
                            spsNoDoc.Formulasi = GetCellValue(firstRow, idxFormulasi); // HARDCODED
                            spsNoDoc.HoseType = GetCellValue(firstRow, hoseTypeCol);
                            spsNoDoc.Dimensi = idxDimensi > 0 ? GetCellValue(firstRow, idxDimensi) : GetCellValue(firstRow, dimensiCol); // Use HARDCODED col 12 (CHS) or col 10 (Legacy)
                            spsNoDoc.Material = GetCellValue(firstRow, materialCol);
                            // Use HARDCODED indices for reliable column detection
                            spsNoDoc.InnerTube = GetCellValue(firstRow, idxInner);
                            spsNoDoc.OuterCover = GetCellValue(firstRow, idxOuter);
                            spsNoDoc.MiddleTube = idxMiddle > 0 ? GetCellValue(firstRow, idxMiddle) : "";
                            // Use Limits
                            spsNoDoc.UseLimitsInner = GetCellValue(firstRow, idxUseLimitsInner);
                            spsNoDoc.UseLimitsOuter = GetCellValue(firstRow, idxUseLimitsOuter);
                            spsNoDoc.UseLimitsMiddle = idxUseLimitsMiddle > 0 ? GetCellValue(firstRow, idxUseLimitsMiddle) : "";
                            // Yarn and tension
                            spsNoDoc.Yarn = idxYarn > 0 ? GetCellValue(firstRow, idxYarn) : "";
                            spsNoDoc.PitchYarn = idxPitchYarn > 0 ? GetCellValue(firstRow, idxPitchYarn) : "";
                            spsNoDoc.TensionYarnInner = idxTensionIn > 0 ? GetCellValue(firstRow, idxTensionIn) : "";
                            spsNoDoc.TensionYarnOuter = idxTensionOut > 0 ? GetCellValue(firstRow, idxTensionOut) : "";
                            // Dies (these store RAW values, will be parsed later)
                            spsNoDoc.Nipple = GetCellValue(firstRow, idxNipple);
                            spsNoDoc.TubeDie = GetCellValue(firstRow, idxTubeDie);
                            spsNoDoc.CoverDie = GetCellValue(firstRow, idxCoverDie);
                            spsNoDoc.MiddleDie = idxMiddleDie > 0 ? GetCellValue(firstRow, idxMiddleDie) : "";
                            spsNoDoc.SpacerDie = idxSpacerDie > 0 ? GetCellValue(firstRow, idxSpacerDie) : "";
                            spsNoDoc.ADistance = idxADistance > 0 ? GetCellValue(firstRow, idxADistance) : "";
                            // Mesh screens
                            spsNoDoc.MeshDim1 = idxMeshDim1 > 0 ? GetCellValue(firstRow, idxMeshDim1) : "";
                            spsNoDoc.MeshScreen1 = idxMeshScreen1 > 0 ? GetCellValue(firstRow, idxMeshScreen1) : "";
                            spsNoDoc.MeshDim2 = idxMeshDim2 > 0 ? GetCellValue(firstRow, idxMeshDim2) : "";
                            spsNoDoc.MeshScreen2 = idxMeshScreen2 > 0 ? GetCellValue(firstRow, idxMeshScreen2) : "";
                            spsNoDoc.MeshDim3 = idxMeshDim3 > 0 ? GetCellValue(firstRow, idxMeshDim3) : "";
                            spsNoDoc.MeshScreen3 = idxMeshScreen3 > 0 ? GetCellValue(firstRow, idxMeshScreen3) : "";
                            // Tolerance fields (RAW values, will be parsed later)
                            spsNoDoc.ToleranceInner = idxToleranceInner > 0 ? GetCellValue(firstRow, idxToleranceInner) : ""; // HARDCODED
                            spsNoDoc.ToleranceOuter = idxToleranceOuter > 0 ? GetCellValue(firstRow, idxToleranceOuter) : ""; // HARDCODED
                        }

                            // ========== PARSE Â± FORMAT FOR TOLERANCE FIELDS ==========
                            _logger.LogInformation($"  Parsing tolerance fields with Â± format...");
                            
                            // DIE/MATERIAL PARAMETERS with Â±
                            AssignFullOrManual(idxNipple,
                                v => spsNoDoc.Nipple_Min = v, v => spsNoDoc.Nipple_Asli = v, v => spsNoDoc.Nipple_Max = v);
                            
                            AssignFullOrManual(idxTubeDie,
                                v => spsNoDoc.TubeDie_Min = v, v => spsNoDoc.TubeDie_Asli = v, v => spsNoDoc.TubeDie_Max = v);
                            
                            AssignFullOrManual(idxCoverDie,
                                v => spsNoDoc.CoverDie_Min = v, v => spsNoDoc.CoverDie_Asli = v, v => spsNoDoc.CoverDie_Max = v);
                            
                            if (idxMiddleDie > 0)
                            {
                                AssignFullOrManual(idxMiddleDie,
                                    v => spsNoDoc.MiddleDie_Min = v, v => spsNoDoc.MiddleDie_Asli = v, v => spsNoDoc.MiddleDie_Max = v);
                            }
                            
                            if (idxSpacerDie > 0)
                            {
                                AssignFullOrManual(idxSpacerDie,
                                    v => spsNoDoc.SpacerDie_Min = v, v => spsNoDoc.SpacerDie_Asli = v, v => spsNoDoc.SpacerDie_Max = v);
                            }
                            
                            if (idxADistance > 0)
                            {
                                AssignFullOrManual(idxADistance,
                                    v => spsNoDoc.ADistance_Min = v, v => spsNoDoc.ADistance_Asli = v, v => spsNoDoc.ADistance_Max = v);
                            }
                            
                            // MESH DIM with Â±
                            if (idxMeshDim1 > 0)
                            {
                                AssignFullOrManual(idxMeshDim1,
                                    v => spsNoDoc.MeshDim1_Min = v, v => spsNoDoc.MeshDim1_Asli = v, v => spsNoDoc.MeshDim1_Max = v);
                            }
                            
                            if (idxMeshDim2 > 0)
                            {
                                AssignFullOrManual(idxMeshDim2,
                                    v => spsNoDoc.MeshDim2_Min = v, v => spsNoDoc.MeshDim2_Asli = v, v => spsNoDoc.MeshDim2_Max = v);
                            }
                            
                            if (idxMeshDim3 > 0)
                            {
                                AssignFullOrManual(idxMeshDim3,
                                    v => spsNoDoc.MeshDim3_Min = v, v => spsNoDoc.MeshDim3_Asli = v, v => spsNoDoc.MeshDim3_Max = v);
                            }
                            
                            // Parse PitchYarn with Â±
                            if (idxPitchYarn > 0)
                            {
                                AssignFullOrManual(idxPitchYarn,
                                    v => spsNoDoc.PitchYarn_Min = v, v => spsNoDoc.PitchYarn_Asli = v, v => spsNoDoc.PitchYarn_Max = v);
                            }
                            
                            // TEMPERATURE/CYLINDER/SPEED PARAMETERS with HARDCODED indices (Â± parsing)
                            // Using hardcoded column indices based on format detection
                            if (idxHeadTemp1 > 0)
                            {
                                AssignFullOrManual(idxHeadTemp1,
                                    v => spsNoDoc.HeadTemp1_Min = v, v => spsNoDoc.HeadTemp1_Asli = v, v => spsNoDoc.HeadTemp1_Max = v);
                            }
                            
                            if (idxHeadTemp2 > 0)
                            {
                                AssignFullOrManual(idxHeadTemp2,
                                    v => spsNoDoc.HeadTemp2_Min = v, v => spsNoDoc.HeadTemp2_Asli = v, v => spsNoDoc.HeadTemp2_Max = v);
                            }
                            
                            if (idxHeadTemp3 > 0)
                            {
                                AssignFullOrManual(idxHeadTemp3,
                                    v => spsNoDoc.HeadTemp3_Min = v, v => spsNoDoc.HeadTemp3_Asli = v, v => spsNoDoc.HeadTemp3_Max = v);
                            }
                            
                            // CYLINDER 1 (Material 1, 2, 3)
                            if (idxCyl11 > 0)
                            {
                                AssignFullOrManual(idxCyl11,
                                    v => spsNoDoc.Cylinder1_1_Min = v, v => spsNoDoc.Cylinder1_1_Asli = v, v => spsNoDoc.Cylinder1_1_Max = v);
                            }
                            
                            if (idxCyl12 > 0)
                            {
                                AssignFullOrManual(idxCyl12,
                                    v => spsNoDoc.Cylinder1_2_Min = v, v => spsNoDoc.Cylinder1_2_Asli = v, v => spsNoDoc.Cylinder1_2_Max = v);
                            }
                            
                            if (idxCyl13 > 0)
                            {
                                AssignFullOrManual(idxCyl13,
                                    v => spsNoDoc.Cylinder1_3_Min = v, v => spsNoDoc.Cylinder1_3_Asli = v, v => spsNoDoc.Cylinder1_3_Max = v);
                            }
                            
                            // CYLINDER 2 (Material 1, 2, 3)
                            if (idxCyl21 > 0)
                            {
                                AssignFullOrManual(idxCyl21,
                                    v => spsNoDoc.Cylinder2_1_Min = v, v => spsNoDoc.Cylinder2_1_Asli = v, v => spsNoDoc.Cylinder2_1_Max = v);
                            }
                            
                            if (idxCyl22 > 0)
                            {
                                AssignFullOrManual(idxCyl22,
                                    v => spsNoDoc.Cylinder2_2_Min = v, v => spsNoDoc.Cylinder2_2_Asli = v, v => spsNoDoc.Cylinder2_2_Max = v);
                            }
                            
                            if (idxCyl23 > 0)
                            {
                                AssignFullOrManual(idxCyl23,
                                    v => spsNoDoc.Cylinder2_3_Min = v, v => spsNoDoc.Cylinder2_3_Asli = v, v => spsNoDoc.Cylinder2_3_Max = v);
                            }
                            
                            // CYLINDER 3 (Material 1, 2, 3)
                            if (idxCyl31 > 0)
                            {
                                AssignFullOrManual(idxCyl31,
                                    v => spsNoDoc.Cylinder3_1_Min = v, v => spsNoDoc.Cylinder3_1_Asli = v, v => spsNoDoc.Cylinder3_1_Max = v);
                            }
                            
                            if (idxCyl32 > 0)
                            {
                                AssignFullOrManual(idxCyl32,
                                    v => spsNoDoc.Cylinder3_2_Min = v, v => spsNoDoc.Cylinder3_2_Asli = v, v => spsNoDoc.Cylinder3_2_Max = v);
                            }
                            
                            if (idxCyl33 > 0)
                            {
                                AssignFullOrManual(idxCyl33,
                                    v => spsNoDoc.Cylinder3_3_Min = v, v => spsNoDoc.Cylinder3_3_Asli = v, v => spsNoDoc.Cylinder3_3_Max = v);
                            }
                            
                            // SCREW TEMPERATURES (Material 1, 2, 3)
                            if (idxScrewTemp1 > 0)
                            {
                                AssignFullOrManual(idxScrewTemp1,
                                    v => spsNoDoc.ScrewTemp1_Min = v, v => spsNoDoc.ScrewTemp1_Asli = v, v => spsNoDoc.ScrewTemp1_Max = v);
                            }
                            
                            if (idxScrewTemp2 > 0)
                            {
                                AssignFullOrManual(idxScrewTemp2,
                                    v => spsNoDoc.ScrewTemp2_Min = v, v => spsNoDoc.ScrewTemp2_Asli = v, v => spsNoDoc.ScrewTemp2_Max = v);
                            }
                            
                            if (idxScrewTemp3 > 0)
                            {
                                AssignFullOrManual(idxScrewTemp3,
                                    v => spsNoDoc.ScrewTemp3_Min = v, v => spsNoDoc.ScrewTemp3_Asli = v, v => spsNoDoc.ScrewTemp3_Max = v);
                            }
                            
                            // SCREW SPEEDS (Material 1, 2, 3)
                            if (idxScrewSpeed1 > 0)
                            {
                                AssignFullOrManual(idxScrewSpeed1,
                                    v => spsNoDoc.ScrewSpeed1_Min = v, v => spsNoDoc.ScrewSpeed1_Asli = v, v => spsNoDoc.ScrewSpeed1_Max = v);
                            }
                            
                            if (idxScrewSpeed2 > 0)
                            {
                                AssignFullOrManual(idxScrewSpeed2,
                                    v => spsNoDoc.ScrewSpeed2_Min = v, v => spsNoDoc.ScrewSpeed2_Asli = v, v => spsNoDoc.ScrewSpeed2_Max = v);
                            }
                            
                            if (idxScrewSpeed3 > 0)
                            {
                                AssignFullOrManual(idxScrewSpeed3,
                                    v => spsNoDoc.ScrewSpeed3_Min = v, v => spsNoDoc.ScrewSpeed3_Asli = v, v => spsNoDoc.ScrewSpeed3_Max = v);
                            }
                            
                            // PRESSURES (Material 1, 2, 3)
                            if (idxPressure1 > 0)
                            {
                                spsNoDoc.Pressure1 = GetCellValue(firstRow, idxPressure1);
                                AssignFullOrManual(idxPressure1,
                                    v => spsNoDoc.Pressure1_Min = v, v => spsNoDoc.Pressure1_Asli = v, v => spsNoDoc.Pressure1_Max = v);
                            }
                            
                            if (idxPressure2 > 0)
                            {
                                spsNoDoc.Pressure2 = GetCellValue(firstRow, idxPressure2);
                                AssignFullOrManual(idxPressure2,
                                    v => spsNoDoc.Pressure2_Min = v, v => spsNoDoc.Pressure2_Asli = v, v => spsNoDoc.Pressure2_Max = v);
                            }
                            
                            if (idxPressure3 > 0)
                            {
                                spsNoDoc.Pressure3 = GetCellValue(firstRow, idxPressure3);
                                AssignFullOrManual(idxPressure3,
                                    v => spsNoDoc.Pressure3_Min = v, v => spsNoDoc.Pressure3_Asli = v, v => spsNoDoc.Pressure3_Max = v);
                            }
                            
                            // FEED (Material 1, 2, 3)
                            int feed1Col = idxFeed1 > 0 ? idxFeed1 : idxFeedRoll1;
                            if (feed1Col > 0)
                            {
                                spsNoDoc.Feed1 = GetCellValue(firstRow, feed1Col);
                                AssignFullOrManual(feed1Col,
                                    v => spsNoDoc.Feed1_Min = v, v => spsNoDoc.Feed1_Asli = v, v => spsNoDoc.Feed1_Max = v);
                            }
                            
                            int feed2Col = idxFeed2 > 0 ? idxFeed2 : idxFeedRoll2;
                            if (feed2Col > 0)
                            {
                                spsNoDoc.Feed2 = GetCellValue(firstRow, feed2Col);
                                AssignFullOrManual(feed2Col,
                                    v => spsNoDoc.Feed2_Min = v, v => spsNoDoc.Feed2_Asli = v, v => spsNoDoc.Feed2_Max = v);
                            }
                            
                            int feed3Col = idxFeed3 > 0 ? idxFeed3 : idxFeedRoll3;
                            if (feed3Col > 0)
                            {
                                spsNoDoc.Feed3 = GetCellValue(firstRow, feed3Col);
                                AssignFullOrManual(feed3Col,
                                    v => spsNoDoc.Feed3_Min = v, v => spsNoDoc.Feed3_Asli = v, v => spsNoDoc.Feed3_Max = v);
                            }

                            // FEED ROLL RATIOS (Material 1, 2, 3)
                            if (idxFeedRoll1 > 0)
                            {
                                spsNoDoc.FeedRollRatio1 = GetCellValue(firstRow, idxFeedRoll1);
                                AssignFullOrManual(idxFeedRoll1,
                                    v => spsNoDoc.FeedRollRatio1_Min = v, v => spsNoDoc.FeedRollRatio1_Asli = v, v => spsNoDoc.FeedRollRatio1_Max = v);
                            }
                            
                            if (idxFeedRoll2 > 0)
                            {
                                spsNoDoc.FeedRollRatio2 = GetCellValue(firstRow, idxFeedRoll2);
                                AssignFullOrManual(idxFeedRoll2,
                                    v => spsNoDoc.FeedRollRatio2_Min = v, v => spsNoDoc.FeedRollRatio2_Asli = v, v => spsNoDoc.FeedRollRatio2_Max = v);
                            }
                            
                            if (idxFeedRoll3 > 0)
                            {
                                spsNoDoc.FeedRollRatio3 = GetCellValue(firstRow, idxFeedRoll3);
                                AssignFullOrManual(idxFeedRoll3,
                                    v => spsNoDoc.FeedRollRatio3_Min = v, v => spsNoDoc.FeedRollRatio3_Asli = v, v => spsNoDoc.FeedRollRatio3_Max = v);
                            }
                            
                            // MACHINE PARAMETERS (Am Meter, Preset, Control, Spiral, Hose Speed)
                            if (idxCurrentValue > 0)
                            {
                                spsNoDoc.CurrentValue = GetCellValue(firstRow, idxCurrentValue);
                            }
                            
                            if (idxAmMeter > 0)
                            {
                                AssignFullOrManual(idxAmMeter,
                                    v => spsNoDoc.AmMeter_Min = v, v => spsNoDoc.AmMeter_Asli = v, v => spsNoDoc.AmMeter_Max = v);
                            }
                            
                            if (idxAmMeter2 > 0)
                            {
                                AssignFullOrManual(idxAmMeter2,
                                    v => spsNoDoc.AmMeter2_Min = v, v => spsNoDoc.AmMeter2_Asli = v, v => spsNoDoc.AmMeter2_Max = v);
                            }
                            
                            if (idxAmMeter3 > 0)
                            {
                                AssignFullOrManual(idxAmMeter3,
                                    v => spsNoDoc.AmMeter3_Min = v, v => spsNoDoc.AmMeter3_Asli = v, v => spsNoDoc.AmMeter3_Max = v);
                            }
                            
                            if (idxPresetValue > 0)
                            {
                                AssignFullOrManual(idxPresetValue,
                                    v => spsNoDoc.PresetValue_Min = v, v => spsNoDoc.PresetValue_Asli = v, v => spsNoDoc.PresetValue_Max = v);
                            }
                            
                            if (idxControlValue > 0)
                            {
                                AssignFullOrManual(idxControlValue,
                                    v => spsNoDoc.ControlValue_Min = v, v => spsNoDoc.ControlValue_Asli = v, v => spsNoDoc.ControlValue_Max = v);
                            }
                            
                            if (idxSpiralPitchSetting > 0)
                            {
                                AssignFullOrManual(idxSpiralPitchSetting,
                                    v => spsNoDoc.SpiralPitchSetting_Min = v, v => spsNoDoc.SpiralPitchSetting_Asli = v, v => spsNoDoc.SpiralPitchSetting_Max = v);
                            }
                            
                            if (idxSpiralPitchDisplay > 0)
                            {
                                AssignFullOrManual(idxSpiralPitchDisplay,
                                    v => spsNoDoc.SpiralPitchDisplay_Min = v, v => spsNoDoc.SpiralPitchDisplay_Asli = v, v => spsNoDoc.SpiralPitchDisplay_Max = v);
                            }
                            
                            if (idxSpiralSpeed > 0)
                            {
                                AssignFullOrManual(idxSpiralSpeed,
                                    v => spsNoDoc.SpiralSpeed_Min = v, v => spsNoDoc.SpiralSpeed_Asli = v, v => spsNoDoc.SpiralSpeed_Max = v);
                            }
                            
                            if (idxHoseSpeed > 0)
                            {
                                AssignFullOrManual(idxHoseSpeed,
                                    v => spsNoDoc.HoseSpeed_Min = v, v => spsNoDoc.HoseSpeed_Asli = v, v => spsNoDoc.HoseSpeed_Max = v);
                            }
                            
                            if (idxUnsmoothSurface > 0)
                            {
                                spsNoDoc.UnsmoothSurface = GetCellValue(firstRow, idxUnsmoothSurface);
                            }
                            
                            // MARKING/CONVEYOR PARAMETERS
                            if (idxMarkingSort > 0)
                            {
                                spsNoDoc.MarkingSort = GetCellValue(firstRow, idxMarkingSort);
                            }
                            
                            if (idxTextMarkingMaterial > 0)
                            {
                                spsNoDoc.TextMarkingMaterial = GetCellValue(firstRow, idxTextMarkingMaterial);
                            }
                            
                            if (idxMarkingColour > 0)
                            {
                                spsNoDoc.MarkingColour = GetCellValue(firstRow, idxMarkingColour);
                            }
                            
                            if (idxChillerWaterTemp > 0)
                            {
                                spsNoDoc.ChillerWaterTemp = GetCellValue(firstRow, idxChillerWaterTemp);
                                AssignFullOrManual(idxChillerWaterTemp,
                                    v => spsNoDoc.ChillerWaterTemp_Min = v, v => spsNoDoc.ChillerWaterTemp_Asli = v, v => spsNoDoc.ChillerWaterTemp_Max = v);
                            }
                            
                            if (idxDancerPosition > 0)
                            {
                                AssignFullOrManual(idxDancerPosition,
                                    v => spsNoDoc.DancerPosition_Min = v, v => spsNoDoc.DancerPosition_Asli = v, v => spsNoDoc.DancerPosition_Max = v);
                            }
                            
                            if (idxCaterpillarGap > 0)
                            {
                                AssignFullOrManual(idxCaterpillarGap,
                                    v => spsNoDoc.CaterpillarGap_Min = v, v => spsNoDoc.CaterpillarGap_Asli = v, v => spsNoDoc.CaterpillarGap_Max = v);
                            }
                            
                            if (idxTakeUpConveyorSpeed > 0)
                            {
                                AssignFullOrManual(idxTakeUpConveyorSpeed,
                                    v => spsNoDoc.TakeUpConveyorSpeed_Min = v, v => spsNoDoc.TakeUpConveyorSpeed_Asli = v, v => spsNoDoc.TakeUpConveyorSpeed_Max = v);
                            }
                            
                            if (idxCoolConveyorSpeed > 0)
                            {
                                AssignFullOrManual(idxCoolConveyorSpeed,
                                    v => spsNoDoc.CoolConveyorSpeed_Min = v, v => spsNoDoc.CoolConveyorSpeed_Asli = v, v => spsNoDoc.CoolConveyorSpeed_Max = v);
                            }
                            
                            if (idxCoolConveyorSpeed2 > 0)
                            {
                                AssignFullOrManual(idxCoolConveyorSpeed2,
                                    v => spsNoDoc.CoolConveyorSpeed2_Min = v, v => spsNoDoc.CoolConveyorSpeed2_Asli = v, v => spsNoDoc.CoolConveyorSpeed2_Max = v);
                            }
                            
                            if (idxConveyorRatio > 0)
                            {
                                AssignFullOrManual(idxConveyorRatio,
                                    v => spsNoDoc.ConveyorRatio_Min = v, v => spsNoDoc.ConveyorRatio_Asli = v, v => spsNoDoc.ConveyorRatio_Max = v);
                            }
                            
                            // TOLERANCE/THICKNESS PARAMETERS - CRITICAL FOR DISPLAY!
                            if (idxToleranceInner > 0)
                            {
                                _logger.LogWarning($"  ðŸ” ToleranceInner from col {idxToleranceInner} (with manual fallback)");
                                AssignFullOrManual(idxToleranceInner,
                                    v => spsNoDoc.ToleranceInner_Min = v, v => spsNoDoc.ToleranceInner_Asli = v, v => spsNoDoc.ToleranceInner_Max = v);
                                _logger.LogWarning($"     âœ“ Parsed â†’ Min={spsNoDoc.ToleranceInner_Min}, Asli={spsNoDoc.ToleranceInner_Asli}, Max={spsNoDoc.ToleranceInner_Max}");
                                
                                if (spsNoDoc.ToleranceInner_Min == null && spsNoDoc.ToleranceInner_Asli == null && spsNoDoc.ToleranceInner_Max == null)
                                {
                                    _logger.LogError($"     âŒ PARSING FAILED! All values are NULL!");
                                }
                            }
                            else
                            {
                                _logger.LogWarning($"  âš ï¸ idxToleranceInner is 0 or negative, skipping parsing");
                            }
                            
                            if (idxToleranceOuter > 0)
                            {
                                _logger.LogWarning($"  ðŸ” ToleranceOuter from col {idxToleranceOuter} (with manual fallback)");
                                AssignFullOrManual(idxToleranceOuter,
                                    v => spsNoDoc.ToleranceOuter_Min = v, v => spsNoDoc.ToleranceOuter_Asli = v, v => spsNoDoc.ToleranceOuter_Max = v);
                                _logger.LogWarning($"     âœ“ Parsed â†’ Min={spsNoDoc.ToleranceOuter_Min}, Asli={spsNoDoc.ToleranceOuter_Asli}, Max={spsNoDoc.ToleranceOuter_Max}");
                                
                                if (spsNoDoc.ToleranceOuter_Min == null && spsNoDoc.ToleranceOuter_Asli == null && spsNoDoc.ToleranceOuter_Max == null)
                                {
                                    _logger.LogError($"     âŒ PARSING FAILED! All values are NULL!");
                                }
                            }
                            else
                            {
                                _logger.LogWarning($"  âš ï¸ idxToleranceOuter is 0 or negative, skipping parsing");
                            }
                            
                            if (idxTebalInner > 0)
                            {
                                _logger.LogWarning($"  ðŸ” TebalInner from col {idxTebalInner} (with manual fallback)");
                                AssignFullOrManual(idxTebalInner,
                                    v => spsNoDoc.TebalInner_Min = v, v => spsNoDoc.TebalInner_Asli = v, v => spsNoDoc.TebalInner_Max = v);
                                _logger.LogWarning($"     âœ“ Parsed â†’ Min={spsNoDoc.TebalInner_Min}, Asli={spsNoDoc.TebalInner_Asli}, Max={spsNoDoc.TebalInner_Max}");
                            }
                            
                            if (idxTebalInnerMiddle > 0)
                            {
                                _logger.LogWarning($"  ðŸ” TebalInnerMiddle from col {idxTebalInnerMiddle} (with manual fallback)");
                                AssignFullOrManual(idxTebalInnerMiddle,
                                    v => spsNoDoc.TebalInnerMiddle_Min = v, v => spsNoDoc.TebalInnerMiddle_Asli = v, v => spsNoDoc.TebalInnerMiddle_Max = v);
                                _logger.LogWarning($"     âœ“ Parsed â†’ Min={spsNoDoc.TebalInnerMiddle_Min}, Asli={spsNoDoc.TebalInnerMiddle_Asli}, Max={spsNoDoc.TebalInnerMiddle_Max}");
                            }
                            
                            if (idxTebalTotal > 0)
                            {
                                _logger.LogWarning($"  ðŸ” TebalTotal from col {idxTebalTotal} (with manual fallback)");
                                AssignFullOrManual(idxTebalTotal,
                                    v => spsNoDoc.TebalTotal_Min = v, v => spsNoDoc.TebalTotal_Asli = v, v => spsNoDoc.TebalTotal_Max = v);
                                _logger.LogWarning($"     âœ“ Parsed â†’ Min={spsNoDoc.TebalTotal_Min}, Asli={spsNoDoc.TebalTotal_Asli}, Max={spsNoDoc.TebalTotal_Max}");
                            }
                            
                            if (idxSelisihTebal > 0)
                            {
                                AssignFullOrManual(idxSelisihTebal,
                                    v => spsNoDoc.SelisihTebal_Min = v, v => spsNoDoc.SelisihTebal_Asli = v, v => spsNoDoc.SelisihTebal_Max = v);
                            }
                            
                            // QUALITY MATRIX (FINAL THICKNESS PARAMETERS)
                            if (idxThickTarget > 0)
                            {
                                spsNoDoc.ThickTarget = GetCellValue(firstRow, idxThickTarget);
                            }
                            
                            if (idxThickTol > 0)
                            {
                                spsNoDoc.ThickTol = GetCellValue(firstRow, idxThickTol);
                            }
                            
                            if (idxThickLCL > 0)
                            {
                                spsNoDoc.ThickLCL = GetCellValue(firstRow, idxThickLCL);
                            }
                            
                            if (idxThickMin > 0)
                            {
                                spsNoDoc.ThickMin = GetCellValue(firstRow, idxThickMin);
                            }
                            
                            if (idxThickUCL > 0)
                            {
                                spsNoDoc.ThickUCL = GetCellValue(firstRow, idxThickUCL);
                            }
                            
                            if (idxThickMax > 0)
                            {
                                spsNoDoc.ThickMax = GetCellValue(firstRow, idxThickMax);
                            }
                            
                            // Quality Matrix: Inner
                            if (idxInnerTarget > 0)
                            {
                                spsNoDoc.InnerTarget = GetCellValue(firstRow, idxInnerTarget);
                            }
                            
                            if (idxInnerTol > 0)
                            {
                                spsNoDoc.InnerTol = GetCellValue(firstRow, idxInnerTol);
                            }
                            
                            if (idxInnerLCL > 0)
                            {
                                spsNoDoc.InnerLCL = GetCellValue(firstRow, idxInnerLCL);
                            }
                            
                            if (idxInnerMin > 0)
                            {
                                spsNoDoc.InnerMin = GetCellValue(firstRow, idxInnerMin);
                            }
                            
                            if (idxInnerUCL > 0)
                            {
                                spsNoDoc.InnerUCL = GetCellValue(firstRow, idxInnerUCL);
                            }
                            
                            if (idxInnerMax > 0)
                            {
                                spsNoDoc.InnerMax = GetCellValue(firstRow, idxInnerMax);
                            }
                            
                            // Quality Matrix: Inner + Middle (CHS 3 Layer only)
                            if (idxInnerMidTarget > 0)
                            {
                                spsNoDoc.InnerMidTarget = GetCellValue(firstRow, idxInnerMidTarget);
                            }
                            
                            if (idxInnerMidTol > 0)
                            {
                                spsNoDoc.InnerMidTol = GetCellValue(firstRow, idxInnerMidTol);
                            }
                            
                            if (idxInnerMidLCL > 0)
                            {
                                spsNoDoc.InnerMidLCL = GetCellValue(firstRow, idxInnerMidLCL);
                            }
                            
                            if (idxInnerMidMin > 0)
                            {
                                spsNoDoc.InnerMidMin = GetCellValue(firstRow, idxInnerMidMin);
                            }
                            
                            if (idxInnerMidUCL > 0)
                            {
                                spsNoDoc.InnerMidUCL = GetCellValue(firstRow, idxInnerMidUCL);
                            }
                            
                            if (idxInnerMidMax > 0)
                            {
                                spsNoDoc.InnerMidMax = GetCellValue(firstRow, idxInnerMidMax);
                            }
                            
                            // Quality Matrix: Total
                            if (idxTotalTarget > 0)
                            {
                                spsNoDoc.TotalTarget = GetCellValue(firstRow, idxTotalTarget);
                            }
                            
                            if (idxTotalTol > 0)
                            {
                                spsNoDoc.TotalTol = GetCellValue(firstRow, idxTotalTol);
                            }
                            
                            if (idxTotalLCL > 0)
                            {
                                spsNoDoc.TotalLCL = GetCellValue(firstRow, idxTotalLCL);
                            }
                            
                            if (idxTotalMin > 0)
                            {
                                spsNoDoc.TotalMin = GetCellValue(firstRow, idxTotalMin);
                            }
                            
                            if (idxTotalUCL > 0)
                            {
                                spsNoDoc.TotalUCL = GetCellValue(firstRow, idxTotalUCL);
                            }
                            
                            if (idxTotalMax > 0)
                            {
                                spsNoDoc.TotalMax = GetCellValue(firstRow, idxTotalMax);
                            }
                            
                            if (idxSpiralSpeed <= 0) {
                                int spiralSpeedCol = FindColumn("SPIRAL SPEED", "SPIRALSPEED");
                                AssignFullOrManual(spiralSpeedCol, 
                                    v => spsNoDoc.SpiralSpeed_Min = v, v => spsNoDoc.SpiralSpeed_Asli = v, v => spsNoDoc.SpiralSpeed_Max = v);
                            }
                            
                            int cuttingSpeedCol = idxCuttingSpeed > 0 ? idxCuttingSpeed : FindColumn("CUTTING SPEED", "CUTTINGSPEED");
                            if (cuttingSpeedCol > 0)
                            {
                                AssignFullOrManual(cuttingSpeedCol, 
                                    v => spsNoDoc.CuttingSpeed_Min = v, v => spsNoDoc.CuttingSpeed_Asli = v, v => spsNoDoc.CuttingSpeed_Max = v);
                            }
                            
                            if (idxTakeUpConveyorSpeed <= 0) {
                                int takeUpSpeedCol = FindColumn("TAKE UP CONVEYOR SPEED", "TAKEUPCONVEYORSPEED");
                                AssignFullOrManual(takeUpSpeedCol, 
                                    v => spsNoDoc.TakeUpConveyorSpeed_Min = v, v => spsNoDoc.TakeUpConveyorSpeed_Asli = v, v => spsNoDoc.TakeUpConveyorSpeed_Max = v);
                            }
                            
                            if (idxChillerWaterTemp <= 0) {
                                int chillerCol = FindColumn("CHILLER WATER TEMP", "CHILLER", "CHILLERWATERTEMP");
                                AssignFullOrManual(chillerCol, 
                                    v => spsNoDoc.ChillerWaterTemp_Min = v, v => spsNoDoc.ChillerWaterTemp_Asli = v, v => spsNoDoc.ChillerWaterTemp_Max = v);
                            }
                            
                            // THICKNESS with ± / MIN / STD / MAX
                            if (idxTebalInner <= 0) {
                                int tebalInnerCol = FindColumn("TEBAL INNER", "TEBALINNER");
                                AssignFullOrManual(tebalInnerCol, 
                                    v => spsNoDoc.TebalInner_Min = v, v => spsNoDoc.TebalInner_Asli = v, v => spsNoDoc.TebalInner_Max = v);
                            }
                            
                            int tebalOuterCol = idxTebalOuter > 0 ? idxTebalOuter : FindColumn("TEBAL OUTER", "TEBALOUTER");
                            if (tebalOuterCol > 0)
                            {
                                AssignFullOrManual(tebalOuterCol, 
                                    v => spsNoDoc.TebalOuter_Min = v, v => spsNoDoc.TebalOuter_Asli = v, v => spsNoDoc.TebalOuter_Max = v);
                            }
                            
                            if (idxTebalTotal <= 0) {
                                int tebalTotalCol = FindColumn("TEBAL TOTAL", "TEBALTOTAL");
                                AssignFullOrManual(tebalTotalCol, 
                                    v => spsNoDoc.TebalTotal_Min = v, v => spsNoDoc.TebalTotal_Asli = v, v => spsNoDoc.TebalTotal_Max = v);
                            }
                            
                            // TOLERANCE Inner/Outer with ± / MIN / STD / MAX
                            if (idxToleranceInner <= 0) {
                                AssignFullOrManual(toleranceInnerCol, 
                                    v => spsNoDoc.ToleranceInner_Min = v, v => spsNoDoc.ToleranceInner_Asli = v, v => spsNoDoc.ToleranceInner_Max = v);
                            }
                            
                            if (idxToleranceOuter <= 0) {
                                AssignFullOrManual(toleranceOuterCol, 
                                    v => spsNoDoc.ToleranceOuter_Min = v, v => spsNoDoc.ToleranceOuter_Asli = v, v => spsNoDoc.ToleranceOuter_Max = v);
                            }
                            
                            // OTHER PARAMETERS with ± / MIN / STD / MAX
                            if (idxPitchYarn <= 0) {
                                int pitchYarnCol = FindColumn("PITCH YARN", "PITCHYARN");
                                AssignFullOrManual(pitchYarnCol, 
                                    v => spsNoDoc.PitchYarn_Min = v, v => spsNoDoc.PitchYarn_Asli = v, v => spsNoDoc.PitchYarn_Max = v);
                            }
                            
                            if (idxCaterpillarGap <= 0) {
                                int caterpillarCol = FindColumn("CATERPILLAR GAP", "CATERPILLARGAP");
                                AssignFullOrManual(caterpillarCol, 
                                    v => spsNoDoc.CaterpillarGap_Min = v, v => spsNoDoc.CaterpillarGap_Asli = v, v => spsNoDoc.CaterpillarGap_Max = v);
                            }
                            
                            if (idxDancerPosition <= 0) {
                                int dancerCol = FindColumn("DANCER POSITION", "DANCERPOSITION");
                                AssignFullOrManual(dancerCol, 
                                    v => spsNoDoc.DancerPosition_Min = v, v => spsNoDoc.DancerPosition_Asli = v, v => spsNoDoc.DancerPosition_Max = v);
                            }
                            
                            int odSensorCol = idxODSensor > 0 ? idxODSensor : FindColumn("OD SENSOR", "ODSENSOR", "OD");
                            if (odSensorCol > 0)
                            {
                                AssignFullOrManual(odSensorCol, 
                                    v => spsNoDoc.OdSensor_Min = v, v => spsNoDoc.OdSensor_Asli = v, v => spsNoDoc.OdSensor_Max = v);
                            }
                            
                            _logger.LogInformation($"  ✓ Tolerance parsing completed");

                            // Log important values for debugging
                            _logger.LogInformation($"  Data from row {firstRow}:");
                            _logger.LogInformation($"    DocumentNumber: '{spsNoDoc.DocumentNumber}'");
                            _logger.LogInformation($"    No: '{spsNoDoc.No}'");
                            _logger.LogInformation($"    Machine: '{spsNoDoc.Machine}'");
                            _logger.LogInformation($"    MachineCode: '{spsNoDoc.MachineCode}'");
                            _logger.LogInformation($"    Customer: '{spsNoDoc.Customer}'");
                            _logger.LogInformation($"    HoseType: '{spsNoDoc.HoseType}'");
                            _logger.LogInformation($"    Formulasi: '{spsNoDoc.Formulasi}'");
                            _logger.LogInformation($"    ToleranceInner parsed: Min={spsNoDoc.ToleranceInner_Min}, Asli={spsNoDoc.ToleranceInner_Asli}, Max={spsNoDoc.ToleranceInner_Max}");
                            _logger.LogInformation($"    ToleranceOuter parsed: Min={spsNoDoc.ToleranceOuter_Min}, Asli={spsNoDoc.ToleranceOuter_Asli}, Max={spsNoDoc.ToleranceOuter_Max}");


                            // Save document (add if new, update if existing)
                            if (isNewDoc)
                            {
                                _context.SpsNoDocs.Add(spsNoDoc);
                                _logger.LogInformation($"✨ Adding NEW document to context");
                            }
                            else
                            {
                                _context.SpsNoDocs.Update(spsNoDoc);
                                _logger.LogInformation($"📝 Updating EXISTING document in context");
                            }
                            
                            await _context.SaveChangesAsync();
                            
                            if (isNewDoc)
                            {
                                importedDocsCount++;
                                _logger.LogInformation($"Created SpsNoDoc: {spsNoDoc.DocumentNumber}");
                            }
                            else
                            {
                                updatedDocsCount++;
                                _logger.LogInformation($"Updated SpsNoDoc: {spsNoDoc.DocumentNumber}");
                            }
                            
                            // Optional backward compatibility: if ItemList column was populated with valid item code
                            if (!string.IsNullOrWhiteSpace(itemListValue))
                            {
                                var items = itemListValue.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(s => s.Trim())
                                    .Where(s => !string.IsNullOrEmpty(s) && !s.StartsWith("UNKNOWN-", StringComparison.OrdinalIgnoreCase))
                                    .ToList();

                                foreach (var item in items)
                                {
                                    bool itemExists = spsNoDoc.ItemLists != null && spsNoDoc.ItemLists.Any(i => i.ItemList == item);
                                    if (!itemExists)
                                    {
                                        _context.SpsItemLists.Add(new SpsItemList
                                        {
                                            DocumentNumber = spsNoDoc.DocumentNumber,
                                            ItemList = item
                                        });
                                        importedCount++;
                                    }
                                }
                                if (importedCount > 0)
                                {
                                    await _context.SaveChangesAsync();
                                }
                            }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"Error processing row {rowNum}");
                        errors.Add($"Row {rowNum}: {ex.Message}");
                        errorCount++;
                    }
                }

                // Auto-refresh dengan redirect
                string successMsg = $"Import berhasil! {importedDocsCount} dokumen baru ditambahkan" +
                    (updatedDocsCount > 0 ? $", {updatedDocsCount} dokumen diperbarui." : ".");
                if (importedCount > 0)
                {
                    successMsg += $" ({importedCount} item list terhubung).";
                }

                TempData["SuccessMessage"] = $"✅ {successMsg}";
                if (errorCount > 0)
                {
                    TempData["WarningMessage"] = $"⚠️ {errorCount} error: {string.Join(", ", errors.Take(3))}";
                }

                await _approvalService.ConsumeApprovalAsync(ApprovalActionType.SpsImportTemplate, fileApprovalKey);
                
                return Json(new
                {
                    success = true,
                    message = successMsg,
                    imported = importedCount,
                    importedDocs = importedDocsCount,
                    updatedDocs = updatedDocsCount,
                    errors = errorCount,
                    errorDetails = errors.Take(10).ToList(),
                    redirect = true
                });

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error importing Excel");
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }


        [HttpGet]
        public IActionResult DownloadEmptyTemplate()
        {
            try
            {
                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Template SPS Master");
                
                // Header - basic SPS fields
                var headers = new[] {
                    "DocumentNumber", "RevisionNumber", "Customer", "HoseType", 
                    "MachineCode", "Machine", "InnerTube", "OuterCover", 
                    "MiddleTube", "Yarn", "Dimensi", "ToleranceInner", "ToleranceOuter"
                };
                
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }
                
                // Style header
                using (var range = worksheet.Cells[1, 1, 1, headers.Length])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                }
                
                // Example data
                worksheet.Cells[2, 1].Value = "DOC-2026-001";
                worksheet.Cells[2, 2].Value = "01";
                worksheet.Cells[2, 3].Value = "PT. Customer";
                worksheet.Cells[2, 4].Value = "Type A";
                worksheet.Cells[2, 5].Value = "DL01";
                worksheet.Cells[2, 6].Value = "Mesin Double Layer 1";
                worksheet.Cells[2, 7].Value = "NBR";
                worksheet.Cells[2, 8].Value = "CR";
                worksheet.Cells[2, 9].Value = "";
                worksheet.Cells[2, 10].Value = "Aramid";
                worksheet.Cells[2, 11].Value = "25.4";
                worksheet.Cells[2, 12].Value = "0.5";
                worksheet.Cells[2, 13].Value = "0.8";
                
                worksheet.Cells.AutoFitColumns();
                
                var stream = new MemoryStream(package.GetAsByteArray());
                var fileName = $"Template_SPS_Master_{DateTime.Now:yyyyMMdd}.xlsx";
                
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating empty template");
                return BadRequest($"Error: {ex.Message}");
            }
        }

        private static string BuildImportApprovalKey(MemoryStream stream)
        {
            var bytes = stream.ToArray();
            var hash = SHA256.HashData(bytes);
            return $"SPS_DOC_IMPORT:{Convert.ToHexString(hash)}";
        }

        /// <summary>
        /// ACTION UTILITY: Fix data lama yang punya multiple ItemLists
        /// Memisahkan 1 record dengan multiple ItemLists menjadi multiple records dengan 1 ItemList each
        /// URL: /SpsMaster/FixMultipleItemLists
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> CleanEmptyDocs()
        {
            try
            {
                var emptyDocs = await _context.SpsNoDocs
                    .Where(x => x.DocumentNumber == null || x.DocumentNumber == "" || x.DocumentNumber == "-" || x.DocumentNumber.StartsWith("ROW-"))
                    .ToListAsync();
                
                int count = emptyDocs.Count;
                if (count > 0)
                {
                    _context.SpsNoDocs.RemoveRange(emptyDocs);
                    await _context.SaveChangesAsync();
                }
                
                return Ok(new { success = true, deleted = count, message = $"Berhasil menghapus {count} data kosong/tidak valid." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        public async Task<IActionResult> FixMultipleItemLists()
        {
            var report = new System.Text.StringBuilder();
            report.AppendLine("<h2>FIX MULTIPLE ITEMLISTS - DATA REPAIR UTILITY</h2>");
            report.AppendLine($"<p><strong>Tanggal:</strong> {DateTime.Now:yyyy-MM-dd HH:mm:ss}</p>");
            report.AppendLine("<hr>");

            try
            {
                // 1. Cari semua SpsNoDoc yang punya lebih dari 1 ItemList
                var problematicRecords = await _context.SpsNoDocs
                    .Include(s => s.ItemLists)
                    .Where(s => s.ItemLists.Count > 1)
                    .ToListAsync();

                report.AppendLine($"<p><strong>Ditemukan {problematicRecords.Count} record dengan multiple ItemLists</strong></p>");
                report.AppendLine("<hr>");

                if (problematicRecords.Count == 0)
                {
                    report.AppendLine("<div style='color: green;'>");
                    report.AppendLine("<h3>âœ“ DATA SUDAH BERSIH!</h3>");
                    report.AppendLine("<p>Tidak ada record dengan multiple ItemLists.</p>");
                    report.AppendLine("</div>");
                    return Content(report.ToString(), "text/html");
                }

                int totalCreated = 0;
                int totalProcessed = 0;

                foreach (var record in problematicRecords)
                {
                    totalProcessed++;
                    var itemLists = record.ItemLists.ToList();
                    
                    report.AppendLine($"<div style='border: 1px solid #ccc; padding: 10px; margin: 10px 0; background: #f5f5f5;'>");
                    report.AppendLine($"<h4>[{totalProcessed}] SpsNoDoc Document: {record.DocumentNumber}</h4>");
                    report.AppendLine($"<p><strong>Original ItemLists:</strong> {string.Join(", ", itemLists.Select(i => i.ItemList))}</p>");

                    // Keep first ItemList in original record
                    var firstItemList = itemLists.First();
                    var otherItemLists = itemLists.Skip(1).ToList();

                    report.AppendLine($"<p style='color: blue;'>â†’ Keeping first ItemList '<strong>{firstItemList.ItemList}</strong>' in record {record.DocumentNumber}</p>");

                    // Create new SpsNoDoc records for other ItemLists
                    foreach (var itemList in otherItemLists)
                    {
                        var newRecord = new SpsNoDoc();
                        
                        // Copy ALL properties using reflection
                        var properties = typeof(SpsNoDoc).GetProperties()
                            .Where(p => p.CanWrite && p.Name != "DocumentNumber" && p.Name != "ItemLists");
                        
                        foreach (var prop in properties)
                        {
                            var value = prop.GetValue(record);
                            prop.SetValue(newRecord, value);
                        }
                        
                        newRecord.DocumentNumber = $"{record.DocumentNumber}-{itemList.ItemList}"; // Ensure unique document number for split records

                        _context.SpsNoDocs.Add(newRecord);
                        await _context.SaveChangesAsync(); // Save to get it created

                        // Add ItemList to new record (keep original ItemList, don't use the one from edited record)
                        _context.SpsItemLists.Add(new SpsItemList
                        {
                            DocumentNumber = newRecord.DocumentNumber,
                            ItemList = itemList.ItemList
                        });

                        totalCreated++;
                        report.AppendLine($"<p style='color: green;'>âœ“ Created new record {newRecord.DocumentNumber} with ItemList '<strong>{itemList.ItemList}</strong>'</p>");
                    }

                    // Remove other ItemLists from original record (keep only first)
                    _context.SpsItemLists.RemoveRange(otherItemLists);
                    await _context.SaveChangesAsync();

                    report.AppendLine($"<p style='color: green;'><strong>âœ“ Success!</strong> Original record kept with 1 ItemList, {otherItemLists.Count} new records created</p>");
                    report.AppendLine("</div>");
                }

                report.AppendLine("<hr>");
                report.AppendLine("<div style='background: #d4edda; padding: 15px; border: 1px solid #c3e6cb; border-radius: 5px;'>");
                report.AppendLine("<h3 style='color: #155724;'>âœ… REPAIR COMPLETED!</h3>");
                report.AppendLine($"<p><strong>Total records processed:</strong> {totalProcessed}</p>");
                report.AppendLine($"<p><strong>Total new records created:</strong> {totalCreated}</p>");
                report.AppendLine("</div>");
                report.AppendLine("<p><a href='/SpsMaster'>â† Kembali ke Index</a></p>");

                _logger.LogInformation($"FixMultipleItemLists completed: {totalProcessed} processed, {totalCreated} created");
            }
            catch (Exception ex)
            {
                report.AppendLine("<hr>");
                report.AppendLine("<div style='background: #f8d7da; padding: 15px; border: 1px solid #f5c6cb; border-radius: 5px;'>");
                report.AppendLine("<h3 style='color: #721c24;'>âŒ ERROR!</h3>");
                report.AppendLine($"<p><strong>Error:</strong> {ex.Message}</p>");
                report.AppendLine($"<pre>{ex.StackTrace}</pre>");
                report.AppendLine("</div>");
                _logger.LogError(ex, "Error fixing multiple ItemLists");
            }

            return Content(report.ToString(), "text/html");
        }
    }
}


