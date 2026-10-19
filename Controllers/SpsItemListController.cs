using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using VelastoProductionSystem.Data;
using VelastoProductionSystem.Models;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml;
using VelastoProductionSystem.Services;
using System.Text.Json;

namespace VelastoProductionSystem.Controllers
{
    public class SpsItemListController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly IApprovalService _approvalService;

        public SpsItemListController(ApplicationDbContext context, IApprovalService approvalService)
        {
            _context = context;
            _approvalService = approvalService;
        }

        private IActionResult RedirectToApprovalRequest(ApprovalActionType actionType, string targetKey, string returnUrl)
        {
            TempData["ErrorMessage"] = "Aksi ini membutuhkan approval admin. Kirim request approval terlebih dahulu.";
            return RedirectToAction("Request", "Approval", new
            {
                actionType,
                targetKey,
                returnUrl
            });
        }

        // GET: SpsItemList
        public async Task<IActionResult> Index(string search, int page = 1, int pageSize = 50)
        {
            if (page < 1) page = 1;
            if (pageSize <= 0) pageSize = 50;

            var query = _context.SpsItemLists.Include(s => s.SpsNoDoc).AsQueryable();

            if (!string.IsNullOrWhiteSpace(search))
            {
                var searchUpper = search.ToUpper();
                var searchNormalized = NormalizeItemListCode(search).ToUpperInvariant();
                query = query.Where(s => 
                    (s.ItemList != null && s.ItemList.ToUpper().Contains(searchUpper)) ||
                    (s.ItemList != null && s.ItemList.Replace("-", "").ToUpper().Contains(searchNormalized)) ||
                    (s.DocumentNumber != null && s.DocumentNumber.ToUpper().Contains(searchUpper))
                );
            }

            var totalCount = await query.CountAsync();
            var totalPages = Math.Max(1, (int)Math.Ceiling(totalCount / (double)pageSize));
            if (page > totalPages) page = totalPages;

            ViewBag.Search = search;
            ViewBag.Page = page;
            ViewBag.PageSize = pageSize;
            ViewBag.TotalPages = totalPages;
            ViewBag.TotalCount = totalCount;

            var rows = await query
                .OrderBy(s => s.ItemList)
                .ThenBy(s => s.DocumentNumber)
                .Skip((page - 1) * pageSize)
                .Take(pageSize)
                .ToListAsync();

            return View(rows);
        }

        [HttpGet]
        public async Task<IActionResult> DownloadExcel(string? search)
        {
            var query = _context.SpsItemLists
                .Include(s => s.SpsNoDoc)
                .AsQueryable();

            if (!string.IsNullOrWhiteSpace(search))
            {
                var searchUpper = search.ToUpper();
                query = query.Where(s =>
                    (s.ItemList != null && s.ItemList.ToUpper().Contains(searchUpper)) ||
                    (s.DocumentNumber != null && s.DocumentNumber.ToUpper().Contains(searchUpper))
                );
            }

            var rows = await query
                .OrderBy(s => s.ItemList)
                .ThenBy(s => s.DocumentNumber)
                .ToListAsync();

            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("SPS Item List");

            var headers = new[] { "No", "ID", "Item List", "No. Document", "Machine", "Hose Type",
                                  "Length Shrink MIN (mm)", "Length Shrink STD (mm)", "Length Shrink MAX (mm)",
                                  "Length Cutting MIN (mm)", "Length Cutting STD (mm)", "Length Cutting MAX (mm)" };
            for (var i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
            }

            using (var range = worksheet.Cells[1, 1, 1, headers.Length])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(44, 62, 80));
                range.Style.Font.Color.SetColor(System.Drawing.Color.White);
            }

            var rowIndex = 2;
            var runningNo = 1;
            foreach (var item in rows)
            {
                worksheet.Cells[rowIndex, 1].Value = runningNo;
                worksheet.Cells[rowIndex, 2].Value = item.Id;
                worksheet.Cells[rowIndex, 3].Value = item.ItemList;
                worksheet.Cells[rowIndex, 4].Value = item.DocumentNumber;
                worksheet.Cells[rowIndex, 5].Value = item.SpsNoDoc?.MachineCode ?? item.SpsNoDoc?.Machine;
                worksheet.Cells[rowIndex, 6].Value = item.SpsNoDoc?.HoseType;
                worksheet.Cells[rowIndex, 7].Value = item.LengthShrinkMin.HasValue ? (object)item.LengthShrinkMin.Value : null;
                worksheet.Cells[rowIndex, 8].Value = item.LengthShrinkStd.HasValue ? (object)item.LengthShrinkStd.Value : null;
                worksheet.Cells[rowIndex, 9].Value = item.LengthShrinkMax.HasValue ? (object)item.LengthShrinkMax.Value : null;
                worksheet.Cells[rowIndex, 10].Value = item.LengthCuttingMin.HasValue ? (object)item.LengthCuttingMin.Value : null;
                worksheet.Cells[rowIndex, 11].Value = item.LengthCuttingStd.HasValue ? (object)item.LengthCuttingStd.Value : null;
                worksheet.Cells[rowIndex, 12].Value = item.LengthCuttingMax.HasValue ? (object)item.LengthCuttingMax.Value : null;
                rowIndex++;
                runningNo++;
            }

            worksheet.Cells.AutoFitColumns();

            var stream = new MemoryStream(package.GetAsByteArray());
            var fileName = $"Sps_ItemList_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        // GET: SpsItemList/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var spsItemList = await _context.SpsItemLists
                .Include(s => s.SpsNoDoc)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (spsItemList == null)
            {
                return NotFound();
            }

            return View(spsItemList);
        }

        // GET: SpsItemList/Create
        public async Task<IActionResult> Create(int? requestId)
        {
            if (requestId.HasValue)
            {
                var currentUser = HttpContext.Session.GetString("UserName") ?? string.Empty;
                var request = await _context.ApprovalRequests
                    .AsNoTracking()
                    .FirstOrDefaultAsync(r => r.Id == requestId.Value
                        && r.ActionType == ApprovalActionType.SpsItemCreate
                        && r.RequesterUserName.ToUpper() == currentUser.ToUpper()
                        && !string.IsNullOrWhiteSpace(r.PayloadJson));

                if (request != null)
                {
                    try
                    {
                        var modelFromRequest = JsonSerializer.Deserialize<SpsItemList>(request.PayloadJson!);
                        if (modelFromRequest != null)
                        {
                            ViewBag.SourceRequestId = request.Id;
                            ViewData["DocumentNumber"] = new SelectList(_context.SpsNoDocs.Where(s => s.IsActive || s.DocumentNumber == modelFromRequest.DocumentNumber), "DocumentNumber", "DocumentNumber", modelFromRequest.DocumentNumber);
                            return View(modelFromRequest);
                        }
                    }
                    catch
                    {
                        // fallback to empty form
                    }
                }
            }

            ViewData["DocumentNumber"] = new SelectList(_context.SpsNoDocs.Where(s => s.IsActive), "DocumentNumber", "DocumentNumber");
            return View();
        }

        // POST: SpsItemList/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,ItemList,DocumentNumber,LengthShrinkMin,LengthShrinkStd,LengthShrinkMax,LengthCuttingMin,LengthCuttingStd,LengthCuttingMax")] SpsItemList spsItemList, bool saveAsDraft = false, int? sourceRequestId = null)
        {
            spsItemList.ItemList = NormalizeItemListCode(spsItemList.ItemList);

            if (_approvalService.IsRequesterRole() && saveAsDraft)
            {
                var itemCode = spsItemList.ItemList?.Trim() ?? string.Empty;
                var docNo = spsItemList.DocumentNumber?.Trim() ?? string.Empty;
                var targetKey = $"{docNo}|{itemCode}";

                await _approvalService.SaveDraftRequestAsync(
                    ApprovalActionType.SpsItemCreate,
                    targetKey,
                    string.IsNullOrWhiteSpace(itemCode) ? "Draft create SPS ItemList" : $"Draft create SPS ItemList {itemCode}",
                    Url.Action(nameof(Create), "SpsItemList") ?? "/SpsItemList/Create",
                    JsonSerializer.Serialize(spsItemList),
                    sourceRequestId);

                TempData["SuccessMessage"] = "Draft Item List tersimpan. Bisa dilanjutkan kapan saja dari Approval Request Saya.";
                return RedirectToAction("MyRequests", "Approval");
            }

            if (string.IsNullOrWhiteSpace(spsItemList.ItemList))
            {
                ModelState.AddModelError("ItemList", "Item List tidak boleh kosong.");
            }

            if (!string.IsNullOrWhiteSpace(spsItemList.DocumentNumber) && !string.IsNullOrWhiteSpace(spsItemList.ItemList))
            {
                var normalizedInput = spsItemList.ItemList.ToUpper();
                var duplicateExists = await _context.SpsItemLists.AnyAsync(i =>
                    i.DocumentNumber == spsItemList.DocumentNumber &&
                    i.ItemList != null &&
                    i.ItemList.Replace("-", "").ToUpper() == normalizedInput);

                if (duplicateExists)
                {
                    ModelState.AddModelError("ItemList", "Item List sudah ada pada No. Document ini.");
                }
            }

            if (ModelState.IsValid)
            {
                if (_approvalService.IsRequesterRole())
                {
                    var targetKey = $"{spsItemList.DocumentNumber}|{spsItemList.ItemList}";

                    var allowed = await _approvalService.HasConsumableApprovalAsync(ApprovalActionType.SpsItemCreate, targetKey);
                    if (!allowed)
                    {
                        await _approvalService.CreateOrReusePendingRequestAsync(
                            ApprovalActionType.SpsItemCreate,
                            targetKey,
                            $"Request create SPS ItemList {spsItemList.ItemList}",
                            Url.Action(nameof(Create), "SpsItemList") ?? "/SpsItemList/Create",
                            JsonSerializer.Serialize(spsItemList),
                            sourceRequestId);

                        TempData["SuccessMessage"] = $"Request data Item List '{spsItemList.ItemList}' terkirim ke SUPERADMIN. Setelah disetujui, data akan otomatis dibuat.";
                        return RedirectToAction("MyRequests", "Approval");
                    }
                }

                _context.Add(spsItemList);
                await _context.SaveChangesAsync();

                await _approvalService.ConsumeApprovalAsync(ApprovalActionType.SpsItemCreate, $"{spsItemList.DocumentNumber}|{spsItemList.ItemList}");
                return RedirectToAction(nameof(Index));
            }
            ViewData["DocumentNumber"] = new SelectList(_context.SpsNoDocs.Where(s => s.IsActive || s.DocumentNumber == spsItemList.DocumentNumber), "DocumentNumber", "DocumentNumber", spsItemList.DocumentNumber);
            return View(spsItemList);
        }

        // GET: SpsItemList/Edit/5
        public async Task<IActionResult> Edit(int? id, int? requestId)
        {
            if (id == null)
            {
                return NotFound();
            }

            var spsItemList = await _context.SpsItemLists.FindAsync(id);
            if (spsItemList == null)
            {
                return NotFound();
            }

            if (requestId.HasValue)
            {
                var currentUser = HttpContext.Session.GetString("UserName") ?? string.Empty;
                var request = await _context.ApprovalRequests
                    .AsNoTracking()
                    .FirstOrDefaultAsync(r => r.Id == requestId.Value
                        && r.ActionType == ApprovalActionType.SpsItemEdit
                        && r.RequesterUserName.ToUpper() == currentUser.ToUpper()
                        && !string.IsNullOrWhiteSpace(r.PayloadJson));

                if (request != null)
                {
                    try
                    {
                        var modelFromRequest = JsonSerializer.Deserialize<SpsItemList>(request.PayloadJson!);
                        if (modelFromRequest != null)
                        {
                            spsItemList.ItemList = modelFromRequest.ItemList;
                            spsItemList.DocumentNumber = modelFromRequest.DocumentNumber;
                            spsItemList.LengthShrinkMin = modelFromRequest.LengthShrinkMin;
                            spsItemList.LengthShrinkStd = modelFromRequest.LengthShrinkStd;
                            spsItemList.LengthShrinkMax = modelFromRequest.LengthShrinkMax;
                            spsItemList.LengthCuttingMin = modelFromRequest.LengthCuttingMin;
                            spsItemList.LengthCuttingStd = modelFromRequest.LengthCuttingStd;
                            spsItemList.LengthCuttingMax = modelFromRequest.LengthCuttingMax;
                            ViewBag.SourceRequestId = request.Id;
                        }
                    }
                    catch
                    {
                        // keep original row
                    }
                }
            }

            ViewData["DocumentNumber"] = new SelectList(_context.SpsNoDocs.Where(s => s.IsActive || s.DocumentNumber == spsItemList.DocumentNumber), "DocumentNumber", "DocumentNumber", spsItemList.DocumentNumber);
            return View(spsItemList);
        }

        // POST: SpsItemList/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("Id,ItemList,DocumentNumber,LengthShrinkMin,LengthShrinkStd,LengthShrinkMax,LengthCuttingMin,LengthCuttingStd,LengthCuttingMax")] SpsItemList spsItemList, bool saveAsDraft = false, int? sourceRequestId = null)
        {
            if (id != spsItemList.Id)
            {
                return NotFound();
            }

            spsItemList.ItemList = NormalizeItemListCode(spsItemList.ItemList);

            if (_approvalService.IsRequesterRole() && saveAsDraft)
            {
                var itemCode = spsItemList.ItemList?.Trim() ?? string.Empty;

                await _approvalService.SaveDraftRequestAsync(
                    ApprovalActionType.SpsItemEdit,
                    spsItemList.Id.ToString(),
                    string.IsNullOrWhiteSpace(itemCode) ? "Draft revisi SPS ItemList" : $"Draft revisi SPS ItemList {itemCode}",
                    Url.Action(nameof(Edit), "SpsItemList", new { id = spsItemList.Id }) ?? $"/SpsItemList/Edit/{spsItemList.Id}",
                    JsonSerializer.Serialize(spsItemList),
                    sourceRequestId);

                TempData["SuccessMessage"] = "Draft revisi Item List tersimpan. Bisa dilanjutkan kapan saja dari Approval Request Saya.";
                return RedirectToAction("MyRequests", "Approval");
            }

            if (string.IsNullOrWhiteSpace(spsItemList.ItemList))
            {
                ModelState.AddModelError("ItemList", "Item List tidak boleh kosong.");
            }

            if (!string.IsNullOrWhiteSpace(spsItemList.DocumentNumber) && !string.IsNullOrWhiteSpace(spsItemList.ItemList))
            {
                var normalizedInput = spsItemList.ItemList.ToUpper();
                var duplicateExists = await _context.SpsItemLists.AnyAsync(i =>
                    i.Id != spsItemList.Id &&
                    i.DocumentNumber == spsItemList.DocumentNumber &&
                    i.ItemList != null &&
                    i.ItemList.Replace("-", "").ToUpper() == normalizedInput);

                if (duplicateExists)
                {
                    ModelState.AddModelError("ItemList", "Item List sudah ada pada No. Document ini.");
                }
            }

            if (ModelState.IsValid)
            {
                try
                {
                    if (_approvalService.IsRequesterRole())
                    {
                        var targetKey = spsItemList.Id.ToString();

                        var allowed = await _approvalService.HasConsumableApprovalAsync(ApprovalActionType.SpsItemEdit, targetKey);
                        if (!allowed)
                        {
                            await _approvalService.CreateOrReusePendingRequestAsync(
                                ApprovalActionType.SpsItemEdit,
                                targetKey,
                                $"Request revisi SPS ItemList {spsItemList.ItemList}",
                                Url.Action(nameof(Edit), "SpsItemList", new { id = spsItemList.Id }) ?? $"/SpsItemList/Edit/{spsItemList.Id}",
                                JsonSerializer.Serialize(spsItemList),
                                sourceRequestId);

                            TempData["SuccessMessage"] = $"Request revisi Item List '{spsItemList.ItemList}' terkirim ke SUPERADMIN. Setelah disetujui, perubahan akan diterapkan.";
                            return RedirectToAction("MyRequests", "Approval");
                        }
                    }

                    _context.Update(spsItemList);
                    await _context.SaveChangesAsync();

                    await _approvalService.ConsumeApprovalAsync(ApprovalActionType.SpsItemEdit, spsItemList.Id.ToString());
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!SpsItemListExists(spsItemList.Id))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            ViewData["DocumentNumber"] = new SelectList(_context.SpsNoDocs.Where(s => s.IsActive || s.DocumentNumber == spsItemList.DocumentNumber), "DocumentNumber", "DocumentNumber", spsItemList.DocumentNumber);
            return View(spsItemList);
        }

        // GET: SpsItemList/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var spsItemList = await _context.SpsItemLists
                .Include(s => s.SpsNoDoc)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (spsItemList == null)
            {
                return NotFound();
            }

            return View(spsItemList);
        }

        // POST: SpsItemList/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var spsItemList = await _context.SpsItemLists.FindAsync(id);
            if (spsItemList != null)
            {
                _context.SpsItemLists.Remove(spsItemList);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        [HttpPost]
        public async Task<IActionResult> DeleteMultiple([FromBody] List<int> ids)
        {
            try
            {
                if (ids == null || !ids.Any())
                {
                    return Json(new { success = false, message = "Tidak ada Item List yang dipilih" });
                }

                var recordsToDelete = await _context.SpsItemLists
                    .Where(s => ids.Contains(s.Id))
                    .ToListAsync();

                if (!recordsToDelete.Any())
                {
                    return Json(new { success = false, message = "Data tidak ditemukan" });
                }

                _context.SpsItemLists.RemoveRange(recordsToDelete);
                await _context.SaveChangesAsync();

                return Json(new { 
                    success = true, 
                    message = $"Berhasil menghapus {recordsToDelete.Count} Item List",
                    deletedCount = recordsToDelete.Count
                });
            }
            catch (Exception ex)
            {
                return Json(new { 
                    success = false, 
                    message = $"Terjadi kesalahan: {ex.Message}" 
                });
            }
        }

        [HttpGet]
        public IActionResult DownloadTemplateItemList()
        {
            try
            {
                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Template ItemList");
                
                worksheet.Cells[1, 1].Value = "ITEM LIST";
                worksheet.Cells[1, 2].Value = "NO.DOC";
                worksheet.Cells[1, 3].Value = "LENGTH SHRINK MIN (mm)";
                worksheet.Cells[1, 4].Value = "LENGTH SHRINK STD (mm)";
                worksheet.Cells[1, 5].Value = "LENGTH SHRINK MAX (mm)";
                worksheet.Cells[1, 6].Value = "LENGTH CUTTING MIN (mm)";
                worksheet.Cells[1, 7].Value = "LENGTH CUTTING STD (mm)";
                worksheet.Cells[1, 8].Value = "LENGTH CUTTING MAX (mm)";
                
                using (var range = worksheet.Cells[1, 1, 1, 8])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189));
                    range.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }
                
                // Contoh data
                worksheet.Cells[2, 1].Value = "NA5030";
                worksheet.Cells[2, 2].Value = "SOP/PROD/HOSE/24/10/038";
                worksheet.Cells[2, 3].Value = 272;
                worksheet.Cells[2, 4].Value = 274;
                worksheet.Cells[2, 5].Value = 276;
                worksheet.Cells[2, 6].Value = 500;
                worksheet.Cells[2, 7].Value = 505;
                worksheet.Cells[2, 8].Value = 510;
                
                worksheet.Cells[4, 1].Value = "Contoh Pengisian:";
                worksheet.Cells[4, 1].Style.Font.Italic = true;
                worksheet.Cells[4, 1].Style.Font.Color.SetColor(System.Drawing.Color.Gray);
                worksheet.Cells[5, 1].Value = "Row 2: NA5030 | SOP/PROD/HOSE/24/10/038 | 272 | 274 | 276 | 500 | 505 | 510";
                worksheet.Cells[5, 1].Style.Font.Size = 9;
                worksheet.Cells[5, 1].Style.Font.Color.SetColor(System.Drawing.Color.Gray);
                worksheet.Cells[6, 1].Value = "Kolom Length Shrink dan Length Cutting boleh dikosongkan jika tidak ada standard.";
                worksheet.Cells[6, 1].Style.Font.Size = 9;
                worksheet.Cells[6, 1].Style.Font.Color.SetColor(System.Drawing.Color.Gray);
                
                worksheet.Cells.AutoFitColumns();
                
                var stream = new MemoryStream(package.GetAsByteArray());
                var fileName = $"Template_ItemList_{DateTime.Now:yyyyMMdd}.xlsx";
                
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error: {ex.Message}");
            }
        }

        [HttpPost]
        public async Task<IActionResult> ImportItemListOnly(IFormFile file)
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
                    var allowed = await _approvalService.HasConsumableApprovalAsync(ApprovalActionType.SpsItemImportTemplate, fileApprovalKey);
                    if (!allowed)
                    {
                        await _approvalService.CreateOrReusePendingRequestAsync(
                            ApprovalActionType.SpsItemImportTemplate,
                            fileApprovalKey,
                            $"Request import Item List file '{file.FileName}'",
                            Url.Action(nameof(Index), "SpsItemList") ?? "/SpsItemList");

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
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                
                if (worksheet == null)
                {
                    return Json(new { success = false, message = "Worksheet tidak ditemukan" });
                }

                int rowCount = worksheet.Dimension?.Rows ?? 0;
                int importedCount = 0;
                int updatedCount = 0;
                int errorCount = 0;
                int duplicateCount = 0;
                int missingDocumentCount = 0;
                int emptyRowCount = 0;
                var errors = new List<string>();
                var missingDocuments = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                var header1 = worksheet.Cells[1, 1].Text?.Trim().ToUpper() ?? "";
                int startRow = 2;
                int itemCol = 0;
                int docCol = 0;
                int shrinkMinCol = 0, shrinkStdCol = 0, shrinkMaxCol = 0;
                int cuttingMinCol = 0, cuttingStdCol = 0, cuttingMaxCol = 0;
                int singleShrinkCol = 0, singleCuttingCol = 0;

                // Scan rows 1 to 4 for header detection
                for (int r = 1; r <= Math.Min(rowCount, 4); r++)
                {
                    bool foundHeaders = false;
                    for (int c = 1; c <= Math.Min(worksheet.Dimension?.Columns ?? 8, 30); c++)
                    {
                        var text = worksheet.Cells[r, c].Text?.Trim().ToUpper() ?? "";
                        if (string.IsNullOrEmpty(text)) continue;

                        var norm = text.Replace(".", "").Replace(" ", "").Replace("_", "").Replace("-", "");

                        if (norm == "ITEM" || norm == "ITEMLIST" || norm == "ITEMCODE" || norm == "KODEITEM") { itemCol = c; foundHeaders = true; }
                        else if (norm == "NODOC" || norm == "DOC" || norm == "DOCUMENT" || norm == "DOCUMENTNUMBER" || norm == "NOMORDOKUMEN" || norm == "NODOKUMEN") { docCol = c; foundHeaders = true; }
                        else if (norm.Contains("SHRINK") && norm.Contains("MIN")) { shrinkMinCol = c; foundHeaders = true; }
                        else if (norm.Contains("SHRINK") && (norm.Contains("STD") || norm.Contains("STN"))) { shrinkStdCol = c; foundHeaders = true; }
                        else if (norm.Contains("SHRINK") && norm.Contains("MAX")) { shrinkMaxCol = c; foundHeaders = true; }
                        else if (norm.Contains("CUT") && norm.Contains("MIN")) { cuttingMinCol = c; foundHeaders = true; }
                        else if (norm.Contains("CUT") && (norm.Contains("STD") || norm.Contains("STN"))) { cuttingStdCol = c; foundHeaders = true; }
                        else if (norm.Contains("CUT") && norm.Contains("MAX")) { cuttingMaxCol = c; foundHeaders = true; }
                        else if (norm.Contains("SHRINK") || norm.Contains("SHRINKAGE")) { singleShrinkCol = c; foundHeaders = true; }
                        else if (norm.Contains("CUT") || norm.Contains("CUTTING")) { singleCuttingCol = c; foundHeaders = true; }
                    }

                    if (foundHeaders && itemCol > 0)
                    {
                        startRow = r + 1;
                        // Skip sub-header rows like "(mm)" or empty item cells with notes
                        while (startRow <= rowCount)
                        {
                            var checkItem = worksheet.Cells[startRow, itemCol].Text?.Trim() ?? "";
                            var checkCol2 = worksheet.Cells[startRow, 2].Text?.Trim() ?? "";
                            if (checkItem.StartsWith("(") || checkCol2.StartsWith("(") || (string.IsNullOrWhiteSpace(checkItem) && !string.IsNullOrWhiteSpace(checkCol2)))
                            {
                                startRow++;
                            }
                            else
                            {
                                break;
                            }
                        }
                        break;
                    }
                }

                // Fallback default columns if no header was detected
                if (itemCol == 0)
                {
                    itemCol = 1;
                    docCol = 2;
                    shrinkMinCol = 3; shrinkStdCol = 4; shrinkMaxCol = 5;
                    cuttingMinCol = 6; cuttingStdCol = 7; cuttingMaxCol = 8;
                }

                float? ParseCleanFloat(string? val)
                {
                    if (string.IsNullOrWhiteSpace(val) || val.Trim() == "-" || val.Trim().Contains("#N/A", StringComparison.OrdinalIgnoreCase)) return null;
                    var cleaned = val.Trim().Replace(" ", "").Replace(",", ".");
                    return float.TryParse(cleaned, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var f) ? f : (float?)null;
                }

                (float? min, float? std, float? max) ParseToleranceOrManual(string? raw, string? minStr = null, string? stdStr = null, string? maxStr = null)
                {
                    float? minVal = ParseCleanFloat(minStr);
                    float? stdVal = ParseCleanFloat(stdStr);
                    float? maxVal = ParseCleanFloat(maxStr);

                    if (!string.IsNullOrWhiteSpace(raw) && !raw.Contains("#N/A", StringComparison.OrdinalIgnoreCase) && raw.Trim() != "-")
                    {
                        var cleaned = raw.Trim().Replace(" ", "").Replace(",", ".");
                        
                        // Check for asymmetric tolerance format: e.g. "10.5+0.1/-0.2"
                        var asymMatch = System.Text.RegularExpressions.Regex.Match(cleaned,
                            @"^([0-9]+(?:\.[0-9]+)?)\s*(?:\()?[\+]+\s*([0-9]+(?:\.[0-9]+)?)\s*[\/\-]+\s*[\-]?\s*([0-9]+(?:\.[0-9]+)?)(?:\))?$",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        if (asymMatch.Success &&
                            float.TryParse(asymMatch.Groups[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float baseVal) &&
                            float.TryParse(asymMatch.Groups[2].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float plusVal) &&
                            float.TryParse(asymMatch.Groups[3].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float minusVal))
                        {
                            return (baseVal - minusVal, baseVal, baseVal + plusVal);
                        }

                        // Check for symmetric tolerance format: e.g. "788±2" or "788+-2"
                        var symMatch = System.Text.RegularExpressions.Regex.Match(cleaned,
                            @"^([0-9]+(?:\.[0-9]+)?)\s*(?:[±\+\-\/]|(?:\+\/\-))\s*([0-9]+(?:\.[0-9]+)?)$",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        if (symMatch.Success &&
                            float.TryParse(symMatch.Groups[1].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float bVal) &&
                            float.TryParse(symMatch.Groups[2].Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float tolVal))
                        {
                            return (bVal - tolVal, bVal, bVal + tolVal);
                        }

                        // If single raw number
                        if (stdVal == null && float.TryParse(cleaned, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float singleVal))
                        {
                            stdVal = singleVal;
                        }
                    }

                    return (minVal, stdVal, maxVal);
                }

                for (int row = startRow; row <= rowCount; row++)
                {
                    try
                    {
                        var itemVal = worksheet.Cells[row, itemCol].Text?.Trim();
                        var docVal = docCol > 0 ? worksheet.Cells[row, docCol].Text?.Trim() : "";

                        if (string.IsNullOrWhiteSpace(itemVal))
                        {
                            emptyRowCount++;
                            continue;
                        }

                        // Ignore header-like or template instruction rows that might repeat
                        if (itemVal.Equals("ITEM", StringComparison.OrdinalIgnoreCase) || 
                            itemVal.Equals("ITEM LIST", StringComparison.OrdinalIgnoreCase) ||
                            itemVal.StartsWith("Contoh", StringComparison.OrdinalIgnoreCase) ||
                            itemVal.StartsWith("Row ", StringComparison.OrdinalIgnoreCase) ||
                            itemVal.StartsWith("Kolom ", StringComparison.OrdinalIgnoreCase) ||
                            itemVal.StartsWith("Catatan", StringComparison.OrdinalIgnoreCase) ||
                            itemVal.StartsWith("Note", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string itemList = NormalizeItemListCode(itemVal);
                        string documentNumber = docVal ?? "";

                        // Parse Shrink
                        var rawShrink = singleShrinkCol > 0 ? worksheet.Cells[row, singleShrinkCol].Text?.Trim() : null;
                        var shrinkMinStr = shrinkMinCol > 0 ? worksheet.Cells[row, shrinkMinCol].Text?.Trim() : null;
                        var shrinkStdStr = shrinkStdCol > 0 ? worksheet.Cells[row, shrinkStdCol].Text?.Trim() : null;
                        var shrinkMaxStr = shrinkMaxCol > 0 ? worksheet.Cells[row, shrinkMaxCol].Text?.Trim() : null;
                        var (shrinkMin, shrinkStd, shrinkMax) = ParseToleranceOrManual(rawShrink ?? shrinkStdStr, shrinkMinStr, shrinkStdStr, shrinkMaxStr);

                        // Parse Cutting
                        var rawCutting = singleCuttingCol > 0 ? worksheet.Cells[row, singleCuttingCol].Text?.Trim() : null;
                        var cuttingMinStr = cuttingMinCol > 0 ? worksheet.Cells[row, cuttingMinCol].Text?.Trim() : null;
                        var cuttingStdStr = cuttingStdCol > 0 ? worksheet.Cells[row, cuttingStdCol].Text?.Trim() : null;
                        var cuttingMaxStr = cuttingMaxCol > 0 ? worksheet.Cells[row, cuttingMaxCol].Text?.Trim() : null;
                        var (cuttingMin, cuttingStd, cuttingMax) = ParseToleranceOrManual(rawCutting ?? cuttingStdStr, cuttingMinStr, cuttingStdStr, cuttingMaxStr);

                        if (string.IsNullOrWhiteSpace(itemList))
                        {
                            errors.Add($"Row {row}: Item List tidak valid setelah normalisasi");
                            errorCount++;
                            continue;
                        }

                        var normalizedItem = itemList.ToUpper();

                        // Case 1: Document number is provided
                        if (!string.IsNullOrWhiteSpace(documentNumber))
                        {
                            var spsNoDoc = await _context.SpsNoDocs.FirstOrDefaultAsync(s => s.DocumentNumber == documentNumber);
                            if (spsNoDoc != null)
                            {
                                var existingItem = await _context.SpsItemLists
                                    .FirstOrDefaultAsync(i => i.DocumentNumber == spsNoDoc.DocumentNumber && i.ItemList != null && i.ItemList.Replace("-", "").ToUpper() == normalizedItem);
                                
                                if (existingItem == null)
                                {
                                    _context.SpsItemLists.Add(new SpsItemList
                                    {
                                        DocumentNumber = spsNoDoc.DocumentNumber,
                                        ItemList = itemList,
                                        LengthShrinkMin = shrinkMin,
                                        LengthShrinkStd = shrinkStd,
                                        LengthShrinkMax = shrinkMax,
                                        LengthCuttingMin = cuttingMin,
                                        LengthCuttingStd = cuttingStd,
                                        LengthCuttingMax = cuttingMax
                                    });
                                    importedCount++;
                                }
                                else
                                {
                                    bool modified = false;
                                    if (shrinkMin.HasValue) { existingItem.LengthShrinkMin = shrinkMin; modified = true; }
                                    if (shrinkStd.HasValue) { existingItem.LengthShrinkStd = shrinkStd; modified = true; }
                                    if (shrinkMax.HasValue) { existingItem.LengthShrinkMax = shrinkMax; modified = true; }
                                    if (cuttingMin.HasValue) { existingItem.LengthCuttingMin = cuttingMin; modified = true; }
                                    if (cuttingStd.HasValue) { existingItem.LengthCuttingStd = cuttingStd; modified = true; }
                                    if (cuttingMax.HasValue) { existingItem.LengthCuttingMax = cuttingMax; modified = true; }
                                    
                                    if (modified)
                                    {
                                        _context.SpsItemLists.Update(existingItem);
                                        updatedCount++;
                                    }
                                    else
                                    {
                                        duplicateCount++;
                                    }
                                }
                            }
                            else
                            {
                                errors.Add($"Row {row}: Document '{documentNumber}' tidak ditemukan");
                                missingDocuments.Add(documentNumber);
                                errorCount++;
                                missingDocumentCount++;
                            }
                        }
                        else
                        {
                            // Case 2: Document number is omitted in file, update existing item(s) by Item code
                            var existingItems = await _context.SpsItemLists
                                .Where(i => i.ItemList != null && i.ItemList.Replace("-", "").ToUpper() == normalizedItem)
                                .ToListAsync();

                            if (existingItems.Any())
                            {
                                bool anyMod = false;
                                foreach (var existingItem in existingItems)
                                {
                                    if (shrinkMin.HasValue) { existingItem.LengthShrinkMin = shrinkMin; anyMod = true; }
                                    if (shrinkStd.HasValue) { existingItem.LengthShrinkStd = shrinkStd; anyMod = true; }
                                    if (shrinkMax.HasValue) { existingItem.LengthShrinkMax = shrinkMax; anyMod = true; }
                                    if (cuttingMin.HasValue) { existingItem.LengthCuttingMin = cuttingMin; anyMod = true; }
                                    if (cuttingStd.HasValue) { existingItem.LengthCuttingStd = cuttingStd; anyMod = true; }
                                    if (cuttingMax.HasValue) { existingItem.LengthCuttingMax = cuttingMax; anyMod = true; }

                                    if (anyMod)
                                    {
                                        _context.SpsItemLists.Update(existingItem);
                                    }
                                }

                                if (anyMod)
                                {
                                    updatedCount += existingItems.Count;
                                }
                                else
                                {
                                    duplicateCount++;
                                }
                            }
                            else
                            {
                                errors.Add($"Row {row}: Item '{itemList}' belum terdaftar pada dokumen manapun (kolom NO.DOC kosong)");
                                errorCount++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Row {row}: {ex.Message}");
                        errorCount++;
                    }
                }

                if (importedCount > 0 || updatedCount > 0)
                {
                    await _context.SaveChangesAsync();
                }

                await _approvalService.ConsumeApprovalAsync(ApprovalActionType.SpsItemImportTemplate, fileApprovalKey);

                if ((importedCount > 0 || updatedCount > 0) && errorCount == 0)
                {
                    return Json(new
                    {
                        success = true,
                        status = "success",
                        message = $"Berhasil import {importedCount} Item List baru" +
                                  (updatedCount > 0 ? $", {updatedCount} data diperbarui." : ".") +
                                  (duplicateCount > 0 ? $" {duplicateCount} data duplikat dilewati." : "") +
                                  (emptyRowCount > 0 ? $" {emptyRowCount} baris kosong diabaikan." : ""),
                        imported = importedCount,
                        updated = updatedCount,
                        duplicates = duplicateCount,
                        errors = errorCount,
                        emptyRows = emptyRowCount
                    });
                }

                if ((importedCount > 0 || updatedCount > 0) && errorCount > 0)
                {
                    return Json(new
                    {
                        success = true,
                        status = "partial",
                        message = $"Import selesai sebagian. Berhasil: {importedCount}" +
                                  (updatedCount > 0 ? $", Diperbarui: {updatedCount}" : "") +
                                  $", Gagal: {errorCount}." +
                                  (missingDocumentCount > 0 ? $" {missingDocumentCount} baris gagal karena No.Doc belum terdaftar." : "") +
                                  (duplicateCount > 0 ? $" {duplicateCount} data duplikat dilewati." : "") +
                                  (emptyRowCount > 0 ? $" {emptyRowCount} baris kosong diabaikan." : ""),
                        imported = importedCount,
                        updated = updatedCount,
                        duplicates = duplicateCount,
                        errors = errorCount,
                        missingDocuments = missingDocuments.Take(10).ToList(),
                        errorDetails = errors.Take(10).ToList(),
                        emptyRows = emptyRowCount
                    });
                }

                if (errorCount > 0)
                {
                    return Json(new
                    {
                        success = false,
                        status = "warning",
                        message = $"Tidak ada data yang tersimpan. {errorCount} baris gagal diproses." +
                                  (missingDocumentCount > 0 ? $" Penyebab utama: No.Doc belum dibuat ({missingDocumentCount} baris)." : ""),
                        imported = importedCount,
                        duplicates = duplicateCount,
                        errors = errorCount,
                        missingDocuments = missingDocuments.Take(10).ToList(),
                        errorDetails = errors.Take(10).ToList(),
                        emptyRows = emptyRowCount
                    });
                }

                return Json(new
                {
                    success = false,
                    status = "warning",
                    message = "Tidak ada data valid untuk diimport. Periksa kembali isi file (Item List dan No.Doc).",
                    imported = 0,
                    duplicates = duplicateCount,
                    errors = 0,
                    emptyRows = emptyRowCount
                });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = $"Error: {ex.Message}" });
            }
        }

        private bool SpsItemListExists(int id)
        {
            return _context.SpsItemLists.Any(e => e.Id == id);
        }

        private static string BuildImportApprovalKey(MemoryStream stream)
        {
            var bytes = stream.ToArray();
            var hash = SHA256.HashData(bytes);
            return $"SPS_ITEM_IMPORT:{Convert.ToHexString(hash)}";
        }

        private static string NormalizeItemListCode(string? rawItemList)
        {
            if (string.IsNullOrWhiteSpace(rawItemList))
            {
                return string.Empty;
            }

            return rawItemList.Trim().Replace("-", string.Empty);
        }
    }
}
