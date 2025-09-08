// Controllers/DocumentController.cs
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

[Authorize]
public class DocumentController : Controller
{
    // 레이아웃 유지용 페이지(iframe 포함)
    [HttpGet]
    public IActionResult Report() => View();

    // 실제 파일을 그대로 스트리밍
    [HttpGet]
    public IActionResult RawReport()
    {
        const string path = @"D:\Development\Web\Client\Report.html";
        if (!System.IO.File.Exists(path))
            return NotFound(@"파일이 없습니다: D:\Development\Web\Client\Report.html");

        return PhysicalFile(path, "text/html; charset=utf-8");
    }
}
