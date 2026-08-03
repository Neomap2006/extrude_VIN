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
                                  "Length Shrink MIN (mm)", "Length Shrink STN (mm)", "Length Shrink MAX (mm)" };
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
        public async Task<IActionResult> Create([Bind("Id,ItemList,DocumentNumber,LengthShrinkMin,LengthShrinkStd,LengthShrinkMax")] SpsItemList spsItemList, bool saveAsDraft = false, int? sourceRequestId = null)
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
        public async Task<IActionResult> Edit(int id, [Bind("Id,ItemList,DocumentNumber,LengthShrinkMin,LengthShrinkStd,LengthShrinkMax")] SpsItemList spsItemList, bool saveAsDraft = false, int? sourceRequestId = null)
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
                worksheet.Cells[1, 4].Value = "LENGTH SHRINK STN (mm)";
                worksheet.Cells[1, 5].Value = "LENGTH SHRINK MAX (mm)";
                
                using (var range = worksheet.Cells[1, 1, 1, 5])
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
                
                worksheet.Cells[4, 1].Value = "Contoh Pengisian:";
                worksheet.Cells[4, 1].Style.Font.Italic = true;
                worksheet.Cells[4, 1].Style.Font.Color.SetColor(System.Drawing.Color.Gray);
                worksheet.Cells[5, 1].Value = "Row 2: HN3040 | SOP/PROD/HOSE/24/10/030 | 121 | 123 | 125";
                worksheet.Cells[5, 1].Style.Font.Size = 9;
                worksheet.Cells[5, 1].Style.Font.Color.SetColor(System.Drawing.Color.Gray);
                worksheet.Cells[6, 1].Value = "Kolom Length Shrink boleh dikosongkan";
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
                int errorCount = 0;
                int duplicateCount = 0;
                int missingDocumentCount = 0;
                int emptyRowCount = 0;
                var errors = new List<string>();
                var missingDocuments = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                var header1 = worksheet.Cells[1, 1].Text?.Trim().ToUpper() ?? "";
                var header2 = worksheet.Cells[1, 2].Text?.Trim().ToUpper() ?? "";
                
                bool isNewFormat = header1.Contains("ITEM") && (header2.Contains("DOC") || header2.Contains("NO"));
                int startRow = 2;

                for (int row = startRow; row <= rowCount; row++)
                {
                    try
                    {
                        var col1Value = worksheet.Cells[row, 1].Text?.Trim();
                        var col2Value = worksheet.Cells[row, 2].Text?.Trim();
                        var col3Value = worksheet.Cells[row, 3].Text?.Trim(); // Length Shrink MIN
                        var col4Value = worksheet.Cells[row, 4].Text?.Trim(); // Length Shrink STN
                        var col5Value = worksheet.Cells[row, 5].Text?.Trim(); // Length Shrink MAX

                        if (string.IsNullOrWhiteSpace(col1Value) || string.IsNullOrWhiteSpace(col2Value))
                        {
                            emptyRowCount++;
                            continue;
                        }

                        string itemList = NormalizeItemListCode(isNewFormat ? col1Value : col2Value);
                        string documentNumber = isNewFormat ? col2Value : col1Value;

                        // Parse Length Shrink values (optional)
                        float? shrinkMin = null, shrinkStd = null, shrinkMax = null;
                        if (!string.IsNullOrWhiteSpace(col3Value) && float.TryParse(col3Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var parsedMin)) shrinkMin = parsedMin;
                        if (!string.IsNullOrWhiteSpace(col4Value) && float.TryParse(col4Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var parsedStd)) shrinkStd = parsedStd;
                        if (!string.IsNullOrWhiteSpace(col5Value) && float.TryParse(col5Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var parsedMax)) shrinkMax = parsedMax;

                        if (string.IsNullOrWhiteSpace(itemList))
                        {
                            errors.Add($"Row {row}: Item List tidak valid setelah normalisasi");
                            errorCount++;
                            continue;
                        }

                        var spsNoDoc = await _context.SpsNoDocs.FirstOrDefaultAsync(s => s.DocumentNumber == documentNumber);
                        if (spsNoDoc != null)
                        {
                            var normalizedItem = itemList.ToUpper();
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
                                    LengthShrinkMax = shrinkMax
                                });
                                importedCount++;
                            }
                            else
                            {
                                duplicateCount++;
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
                    catch (Exception ex)
                    {
                        errors.Add($"Row {row}: {ex.Message}");
                        errorCount++;
                    }
                }

                if (importedCount > 0)
                {
                    await _context.SaveChangesAsync();
                }

                await _approvalService.ConsumeApprovalAsync(ApprovalActionType.SpsItemImportTemplate, fileApprovalKey);

                if (importedCount > 0 && errorCount == 0)
                {
                    return Json(new
                    {
                        success = true,
                        status = "success",
                        message = $"Berhasil import {importedCount} Item List baru." +
                                  (duplicateCount > 0 ? $" {duplicateCount} data duplikat dilewati." : "") +
                                  (emptyRowCount > 0 ? $" {emptyRowCount} baris kosong diabaikan." : ""),
                        imported = importedCount,
                        duplicates = duplicateCount,
                        errors = errorCount,
                        emptyRows = emptyRowCount
                    });
                }

                if (importedCount > 0 && errorCount > 0)
                {
                    return Json(new
                    {
                        success = true,
                        status = "partial",
                        message = $"Import selesai sebagian. Berhasil: {importedCount}, Gagal: {errorCount}." +
                                  (missingDocumentCount > 0 ? $" {missingDocumentCount} baris gagal karena No.Doc belum terdaftar." : "") +
                                  (duplicateCount > 0 ? $" {duplicateCount} data duplikat dilewati." : "") +
                                  (emptyRowCount > 0 ? $" {emptyRowCount} baris kosong diabaikan." : ""),
                        imported = importedCount,
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
