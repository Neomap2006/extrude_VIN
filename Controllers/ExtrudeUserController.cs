using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;
using VelastoProductionSystem.Data;
using VelastoProductionSystem.Models;

namespace VelastoProductionSystem.Controllers
{
    public class ExtrudeUserController : Controller
    {
        private readonly ElwpDbContext _context;
        private readonly ILogger<ExtrudeUserController> _logger;

        public ExtrudeUserController(ElwpDbContext context, ILogger<ExtrudeUserController> logger)
        {
            _context = context;
            _logger = logger;
        }

        // GET: ExtrudeUser
        public async Task<IActionResult> Index()
        {
            try
            {
                // Filter hanya area Extrude (ID 1)
                var users = await _context.ElwpUsers
                    .Include(u => u.Area)
                    .Where(u => u.AreaId == 1)
                    .OrderBy(u => u.FullName)
                    .ToListAsync();

                return View(users);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error getting user list");
                ViewBag.ErrorMessage = "Gagal memuat data user: " + ex.Message;
                return View(new List<ElwpUser>());
            }
        }

        // GET: ExtrudeUser/Create
        public IActionResult Create()
        {
            return View(new ElwpUser { AreaId = 1, IsActive = true, PlantId = 1, Role = "Operator" });
        }

        // POST: ExtrudeUser/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(ElwpUser user)
        {
            // Sanitasi input dasar
            user.Username = user.Username?.Trim() ?? string.Empty;
            user.FullName = user.FullName?.Trim() ?? string.Empty;
            user.NPK = string.IsNullOrWhiteSpace(user.NPK) ? null : user.NPK.Trim();
            user.Email = string.IsNullOrWhiteSpace(user.Email) ? null : user.Email.Trim();
            user.Role = string.IsNullOrWhiteSpace(user.Role) ? "Operator" : user.Role.Trim();
            user.PlantId = user.PlantId > 0 ? user.PlantId : 1;
            user.AreaId = (user.AreaId.HasValue && user.AreaId.Value > 0) ? user.AreaId.Value : 1;

            if (string.IsNullOrWhiteSpace(user.Username))
            {
                ModelState.AddModelError("Username", "Username wajib diisi.");
            }

            if (string.IsNullOrWhiteSpace(user.FullName))
            {
                ModelState.AddModelError("FullName", "Nama Lengkap wajib diisi.");
            }

            // Validasi duplikasi Username
            if (!string.IsNullOrWhiteSpace(user.Username) && await _context.ElwpUsers.AnyAsync(u => u.Username.ToLower() == user.Username.ToLower()))
            {
                ModelState.AddModelError("Username", $"Username '{user.Username}' sudah terdaftar. Gunakan username lain.");
            }

            // Validasi duplikasi Email (jika diisi)
            if (!string.IsNullOrWhiteSpace(user.Email) && await _context.ElwpUsers.AnyAsync(u => u.Email != null && u.Email.ToLower() == user.Email.ToLower()))
            {
                ModelState.AddModelError("Email", $"Email '{user.Email}' sudah digunakan.");
            }

            if (ModelState.IsValid)
            {
                try
                {
                    var hasher = new PasswordHasher<ElwpUser>();
                    var plainPassword = !string.IsNullOrWhiteSpace(user.PasswordHash)
                        ? user.PasswordHash
                        : "extrude123!"; // Default standard operator password

                    user.PasswordHash = hasher.HashPassword(user, plainPassword);
                    user.CreatedAt = DateTime.Now;

                    _context.Add(user);
                    await _context.SaveChangesAsync();

                    TempData["SuccessMessage"] = $"User Extrude '{user.FullName}' berhasil ditambahkan!";
                    return RedirectToAction(nameof(Index));
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error creating user");
                    ModelState.AddModelError("", "Gagal menyimpan user ke database: " + (ex.InnerException?.Message ?? ex.Message));
                }
            }

            return View(user);
        }

        // GET: ExtrudeUser/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null) return NotFound();

            var user = await _context.ElwpUsers.FindAsync(id);
            if (user == null || user.AreaId != 1) return NotFound();

            return View(user);
        }

        // POST: ExtrudeUser/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, ElwpUser user)
        {
            if (id != user.Id) return NotFound();

            // Sanitasi input dasar
            user.Username = user.Username?.Trim() ?? string.Empty;
            user.FullName = user.FullName?.Trim() ?? string.Empty;
            user.NPK = string.IsNullOrWhiteSpace(user.NPK) ? null : user.NPK.Trim();
            user.Email = string.IsNullOrWhiteSpace(user.Email) ? null : user.Email.Trim();
            user.Role = string.IsNullOrWhiteSpace(user.Role) ? "Operator" : user.Role.Trim();

            if (string.IsNullOrWhiteSpace(user.Username))
            {
                ModelState.AddModelError("Username", "Username wajib diisi.");
            }

            if (string.IsNullOrWhiteSpace(user.FullName))
            {
                ModelState.AddModelError("FullName", "Nama Lengkap wajib diisi.");
            }

            // Validasi duplikasi Username
            if (!string.IsNullOrWhiteSpace(user.Username) && await _context.ElwpUsers.AnyAsync(u => u.Id != id && u.Username.ToLower() == user.Username.ToLower()))
            {
                ModelState.AddModelError("Username", $"Username '{user.Username}' sudah digunakan oleh user lain.");
            }

            // Validasi duplikasi Email
            if (!string.IsNullOrWhiteSpace(user.Email) && await _context.ElwpUsers.AnyAsync(u => u.Id != id && u.Email != null && u.Email.ToLower() == user.Email.ToLower()))
            {
                ModelState.AddModelError("Email", $"Email '{user.Email}' sudah digunakan oleh user lain.");
            }

            if (ModelState.IsValid)
            {
                try
                {
                    var existingUser = await _context.ElwpUsers.FindAsync(id);
                    if (existingUser == null) return NotFound();

                    existingUser.FullName = user.FullName;
                    existingUser.Username = user.Username;
                    existingUser.NPK = user.NPK;
                    existingUser.Email = user.Email;
                    existingUser.Role = user.Role;
                    existingUser.IsActive = user.IsActive;

                    if (!string.IsNullOrWhiteSpace(user.PasswordHash))
                    {
                        var hasher = new PasswordHasher<ElwpUser>();
                        existingUser.PasswordHash = hasher.HashPassword(existingUser, user.PasswordHash);
                    }

                    _context.Update(existingUser);
                    await _context.SaveChangesAsync();

                    TempData["SuccessMessage"] = $"Data user '{existingUser.FullName}' berhasil diperbarui!";
                    return RedirectToAction(nameof(Index));
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!UserExists(user.Id)) return NotFound();
                    else throw;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error updating user");
                    ModelState.AddModelError("", "Gagal memperbarui data user: " + (ex.InnerException?.Message ?? ex.Message));
                }
            }

            return View(user);
        }

        // POST: ExtrudeUser/Delete/5
        [HttpPost]
        public async Task<IActionResult> Delete(int id)
        {
            try
            {
                var user = await _context.ElwpUsers.FindAsync(id);
                if (user != null && user.AreaId == 1)
                {
                    _context.ElwpUsers.Remove(user);
                    await _context.SaveChangesAsync();
                    return Json(new { success = true, message = "User berhasil dihapus!" });
                }
                return Json(new { success = false, message = "User tidak ditemukan!" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error deleting user");
                return Json(new { success = false, message = "Gagal menghapus user: " + (ex.InnerException?.Message ?? ex.Message) });
            }
        }

        private bool UserExists(int id)
        {
            return _context.ElwpUsers.Any(e => e.Id == id);
        }
    }
}
