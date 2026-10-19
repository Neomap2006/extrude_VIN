using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace VelastoProductionSystem.Models
{
    public class ElwpUser
    {
        public int Id { get; set; }

        [Required(ErrorMessage = "Username wajib diisi")]
        [StringLength(100)]
        public string Username { get; set; } = string.Empty;

        [StringLength(255)]
        public string? PasswordHash { get; set; }

        [Required(ErrorMessage = "Nama Lengkap wajib diisi")]
        [StringLength(100)]
        public string FullName { get; set; } = string.Empty;

        [EmailAddress(ErrorMessage = "Format email tidak valid")]
        [StringLength(255)]
        public string? Email { get; set; }

        [StringLength(50)]
        public string? NPK { get; set; }

        public int PlantId { get; set; } = 1;
        public int? AreaId { get; set; } = 1;

        [StringLength(50)]
        public string? Role { get; set; } = "Operator";

        public bool IsActive { get; set; } = true;

        public DateTime? CreatedAt { get; set; }
        public DateTime? LastLoginAt { get; set; }

        // Navigation property (optional for join)
        [ForeignKey("AreaId")]
        public virtual ElwpArea? Area { get; set; }
    }
}
