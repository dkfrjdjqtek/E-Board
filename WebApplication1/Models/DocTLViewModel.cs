using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;

namespace WebApplication1.Models
{
    public class DocTLViewModel
    {
        public string? SelectedCompCd { get; set; }
        public int? SelectedDepartmentId { get; set; }
        public string? SelectedDocumentCode { get; set; }

        public List<SelectListItem> CompOptions { get; set; } = new();
        public List<SelectListItem> DepartmentOptions { get; set; } = new();
        public List<SelectListItem> DocumentOptions { get; set; } = new();
    }
}
