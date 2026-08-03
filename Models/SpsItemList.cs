using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace VelastoProductionSystem.Models
{
    public class SpsItemList
    {
        public int Id { get; set; }

        [Required]
        [Display(Name = "Item List")]
        public string ItemList { get; set; } = string.Empty;

        [Required]
        [Display(Name = "No. Document")]
        public string DocumentNumber { get; set; } = string.Empty;

        [ForeignKey("DocumentNumber")]
        public SpsNoDoc? SpsNoDoc { get; set; }

        public float? LengthShrinkMin { get; set; }
        
        [Display(Name = "Length Shrink (mm)")]
        public float? LengthShrinkStd { get; set; }
        
        public float? LengthShrinkMax { get; set; }
    }
}
