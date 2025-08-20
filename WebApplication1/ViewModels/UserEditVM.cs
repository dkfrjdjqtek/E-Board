// ViewModels/UserEditVM.cs
using System.Collections.Generic;
using WebApplication1.ViewModels;


namespace WebApplication1.ViewModels
{
    public record UserRowVM
    {
        public string UserId { get; init; } = default!;
        public string Email { get; init; } = default!;
        public string? DisplayName { get; init; }
        public string Department { get; init; } = "";
        public string Position { get; init; } = "";
    }

    public class UserEditVM
    {
        public string UserId { get; set; } = default!;
        public string Email { get; set; } = default!;
        public string? DisplayName { get; set; }
        public string CompCd { get; set; } = "0001";
        public int? DepartmentId { get; set; }
        public int? PositionId { get; set; }

        // 드롭다운 데이터
        public List<KeyValuePair<int, string>> Departments { get; set; } = new();
        public List<KeyValuePair<int, string>> Positions { get; set; } = new();
    }
}
