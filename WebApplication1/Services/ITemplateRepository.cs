// 2025.10.15 Added: 템플릿/디스크립터 실제 저장 연동을 위한 리포지토리 인터페이스 및 메모리 스텁
using System.Threading.Tasks;

namespace WebApplication1.Services
{
    // 2025.10.15 Added: 템플릿 메타 모델(리소스 키/다국어 대응은 상위 레이어에서 처리)
    public sealed class TemplateMeta
    {
        public string TemplateCode { get; init; } = string.Empty;   // 템플릿 코드(키)
        public string TemplateTitle { get; init; } = "";            // 템플릿 제목(다국어 키 가능)
        public string DescriptorJson { get; init; } = "{}";         // 디스크립터 JSON
        public string PreviewJson { get; init; } = "{}";            // 프리뷰 JSON(미리보기 데이터)
        public string? Version { get; init; }                        // 디스크립터 버전(옵션)
    }

    // 2025.10.15 Added: 저장소 추상화(향후 DB/파일/캐시 교체 용이)
    public interface ITemplateRepository
    {
        // 2025.10.15 Added: 템플릿 코드로 메타 조회(없으면 null)
        Task<TemplateMeta?> GetAsync(string templateCode);
    }

    // 2025.10.15 Added: 메모리 구현체(운영 전 DB/파일형 구현으로 교체)
    // 주의: 데모/개발용이며 앱 시작 시 DI 등록 필요(Program.cs/Startup.cs)
    public sealed class InMemoryTemplateRepository : ITemplateRepository
    {
        // 2025.10.15 Added: 간단한 샘플 1건 등록. 필요 시 자유롭게 추가
        private static readonly System.Collections.Generic.Dictionary<string, TemplateMeta> _data =
            new(System.StringComparer.OrdinalIgnoreCase)
            {
                ["HR_LEAVE_V1"] = new TemplateMeta
                {
                    TemplateCode = "HR_LEAVE_V1",
                    TemplateTitle = "DOC_Tmpl_HR_Leave", // 다국어 리소스 키 사용 예
                    DescriptorJson = @"{
                        ""version"": ""1.0"",
                        ""inputs"": [
                            { ""key"": ""empName"",  ""type"": ""text"",   ""required"": true },
                            { ""key"": ""fromDate"", ""type"": ""date"",   ""required"": true, ""a1"": ""B5"" },
                            { ""key"": ""toDate"",   ""type"": ""date"",   ""required"": true, ""a1"": ""B6"" },
                            { ""key"": ""reason"",   ""type"": ""text"",   ""required"": false }
                        ],
                        ""approvals"": [
                            { ""roleKey"": ""Approver_DeptHead"", ""approverType"": ""Person"", ""required"": true },
                            { ""roleKey"": ""Approver_HR"",       ""approverType"": ""Person"", ""required"": true }
                        ]
                    }",
                    PreviewJson = @"{
                        ""sheets"": [{ ""name"": ""Preview"", ""cells"": { ""B5"": ""2025-10-15"", ""B6"": ""2025-10-16"" } }]
                    }",
                    Version = "1.0"
                }
            };

        // 2025.10.15 Added: 조회 구현
        public Task<TemplateMeta?> GetAsync(string templateCode)
        {
            _data.TryGetValue(templateCode ?? string.Empty, out var meta);
            return Task.FromResult(meta);
        }
    }
}
