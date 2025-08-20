using System.ComponentModel.DataAnnotations;

namespace WebApplication1.Models
{
    public class CompMaster
    {
        [Key, MaxLength(10)]
        public string CompCd { get; set; } = default!;

        [Required, MaxLength(100)]
        public string Name { get; set; } = default!;

        public bool IsActive { get; set; } = true;
    }
}