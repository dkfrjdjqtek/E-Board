// 2025.09.09
using Microsoft.AspNetCore.Mvc.Rendering;

namespace WebApplication1.Models
{
    // 2025.09.09
    public class DocTLViewModel
    {
        public List<SelectListItem> CompOptions { get; set; } = new();
        public List<SelectListItem> DepartmentOptions { get; set; } = new();
        public List<SelectListItem> DocumentOptions { get; set; } = new();
    }
}
