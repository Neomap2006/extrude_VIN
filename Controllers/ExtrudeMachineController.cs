using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using VelastoProductionSystem.Data;
using VelastoProductionSystem.Models;

namespace VelastoProductionSystem.Controllers
{
    public class ExtrudeMachineController : Controller
    {
        private readonly ElwpDbContext _context;
        private readonly ILogger<ExtrudeMachineController> _logger;

        public ExtrudeMachineController(ElwpDbContext context, ILogger<ExtrudeMachineController> logger)
        {
            _context = context;
            _logger = logger;
        }

        // GET: ExtrudeMachine
        public async Task<IActionResult> Index()
        {
            try
            {
                // Filter area Extrude (ID 1)
                var machines = await _context.ElwpMachines
                    .Include(m => m.Area)
                    .Where(m => m.AreaId == 1)
                    .OrderBy(m => m.AreaId)
                    .ThenBy(m => m.KodeMesin)
                    .ToListAsync();
                
                return View(machines);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error loading machine list");
                ViewBag.ErrorMessage = "Gagal memuat data mesin: " + ex.Message;
                return View(new List<ElwpMachine>());
            }
        }

        // GET: ExtrudeMachine/Create
        public IActionResult Create()
        {
            return View(new ElwpMachine { AreaId = 1, PlantId = 1, IsActive = true });
        }

        // POST: ExtrudeMachine/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(ElwpMachine machine)
        {
            machine.KodeMesin = machine.KodeMesin?.Trim().ToUpperInvariant() ?? string.Empty;
            machine.NamaMesin = machine.NamaMesin?.Trim() ?? string.Empty;
            machine.AreaId = machine.AreaId > 0 ? machine.AreaId : 1;
            machine.PlantId = machine.PlantId > 0 ? machine.PlantId : 1;

            if (string.IsNullOrWhiteSpace(machine.KodeMesin))
            {
                ModelState.AddModelError("KodeMesin", "Kode Mesin wajib diisi.");
            }

            if (string.IsNullOrWhiteSpace(machine.NamaMesin))
            {
                ModelState.AddModelError("NamaMesin", "Nama Mesin wajib diisi.");
            }

            if (!string.IsNullOrWhiteSpace(machine.KodeMesin) && await _context.ElwpMachines.AnyAsync(m => m.KodeMesin == machine.KodeMesin))
            {
                ModelState.AddModelError("KodeMesin", $"Kode Mesin '{machine.KodeMesin}' sudah terdaftar.");
            }

            if (ModelState.IsValid)
            {
                try
                {
                    machine.CreatedAt = DateTime.Now;
                    _context.Add(machine);
                    await _context.SaveChangesAsync();
                    TempData["SuccessMessage"] = $"Mesin '{machine.KodeMesin}' berhasil ditambahkan!";
                    return RedirectToAction(nameof(Index));
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error creating machine");
                    ModelState.AddModelError("", "Gagal menyimpan mesin: " + (ex.InnerException?.Message ?? ex.Message));
                }
            }
            return View(machine);
        }

        // GET: ExtrudeMachine/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null) return NotFound();

            var machine = await _context.ElwpMachines.FindAsync(id);
            if (machine == null) return NotFound();

            return View(machine);
        }

        // POST: ExtrudeMachine/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, ElwpMachine machine)
        {
            if (id != machine.Id) return NotFound();

            machine.KodeMesin = machine.KodeMesin?.Trim().ToUpperInvariant() ?? string.Empty;
            machine.NamaMesin = machine.NamaMesin?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(machine.KodeMesin))
            {
                ModelState.AddModelError("KodeMesin", "Kode Mesin wajib diisi.");
            }

            if (string.IsNullOrWhiteSpace(machine.NamaMesin))
            {
                ModelState.AddModelError("NamaMesin", "Nama Mesin wajib diisi.");
            }

            if (!string.IsNullOrWhiteSpace(machine.KodeMesin) && await _context.ElwpMachines.AnyAsync(m => m.Id != id && m.KodeMesin == machine.KodeMesin))
            {
                ModelState.AddModelError("KodeMesin", $"Kode Mesin '{machine.KodeMesin}' sudah digunakan oleh mesin lain.");
            }

            if (ModelState.IsValid)
            {
                try
                {
                    var existing = await _context.ElwpMachines.FindAsync(id);
                    if (existing == null) return NotFound();

                    existing.KodeMesin = machine.KodeMesin;
                    existing.NamaMesin = machine.NamaMesin;
                    existing.Keterangan = machine.Keterangan;
                    existing.AreaId = machine.AreaId > 0 ? machine.AreaId : 1;
                    existing.IsActive = machine.IsActive;
                    existing.Kapasitas = machine.Kapasitas;
                    existing.RequiredManPower = machine.RequiredManPower;
                    existing.UpdatedAt = DateTime.Now;

                    _context.Update(existing);
                    await _context.SaveChangesAsync();
                    TempData["SuccessMessage"] = $"Data mesin '{existing.KodeMesin}' berhasil diperbarui!";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!MachineExists(machine.Id)) return NotFound();
                    else throw;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error updating machine");
                    ModelState.AddModelError("", "Gagal memperbarui mesin: " + (ex.InnerException?.Message ?? ex.Message));
                }
            }
            return View(machine);
        }

        // POST: ExtrudeMachine/Delete/5
        [HttpPost]
        public async Task<IActionResult> Delete(int id)
        {
            try
            {
                var machine = await _context.ElwpMachines.FindAsync(id);
                if (machine != null)
                {
                    _context.ElwpMachines.Remove(machine);
                    await _context.SaveChangesAsync();
                    return Json(new { success = true, message = "Mesin berhasil dihapus!" });
                }
                return Json(new { success = false, message = "Mesin tidak ditemukan!" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting machine");
                return Json(new { success = false, message = "Gagal menghapus mesin: " + (ex.InnerException?.Message ?? ex.Message) });
            }
        }

        private bool MachineExists(int id)
        {
            return _context.ElwpMachines.Any(e => e.Id == id);
        }
    }
}
