// 2026.05.07 Changed: 일반 사용자 프로필 저장 시 사업장을 기존 프로필 사업장으로 고정하고 부서와 직급을 CompCd 기준으로 조회 및 검증
using System.ComponentModel.DataAnnotations;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Localization;
using System.Globalization;
using System.Text.Encodings.Web;
using WebApplication1.Data;
using WebApplication1.Models;
using NuGet.Protocol.Plugins;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Hosting;    // 2025.12.10 Changed: 서명 이미지 저장용 컨텐트 루트 사용
using Microsoft.AspNetCore.Http;       // 2025.12.02 Added: 서명 파일 바인딩용
using System.IO;                       // 2025.12.02 Added: 파일 저장 IO 사용
//using System.Drawing;                  // 2025.12.02 Added: 서명 이미지 리사이즈
//using System.Drawing.Drawing2D;        // 2025.12.02 Added: 고품질 리사이즈 옵션
//using System.Drawing.Imaging;          // 2025.12.02 Added: PNG 저장 포맷
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SixLabors.ImageSharp.Formats.Png;

namespace WebApplication1.Areas.Identity.Pages.Account
{
    [Authorize]
    public class ChangeProfileModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly ApplicationDbContext _db;
        private readonly IEmailSender _emailSender;
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IOptionsMonitor<DataProtectionTokenProviderOptions> _tokenOpts;
        // 2025.12.10 Changed: 서명 이미지 저장 경로 기준을 WebRoot → ContentRoot 로 변경
        private readonly IWebHostEnvironment _env;

        // 2025.12.02 Added: 서명 이미지 기본 경로 상수
        // 요청 URL 은 그대로 /images/signatures/상대경로 를 사용 (향후 전용 액션으로 매핑 예정)
        private const string SignatureRequestBasePath = "/images/signatures";
        // 2025.12.10 Changed: 실제 저장 루트를 ContentRoot\App_Data\Signatures 로 변경
        private const string SignatureRootRelativeFolder = "App_Data/Signatures";

        // 2025.12.02 Added: 서명 이미지 미리보기 URL 바인딩용
        public string? SignatureImageUrl { get; set; }

        private static string Ui2() => CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;

        public ChangeProfileModel(
            UserManager<ApplicationUser> userManager,
            SignInManager<ApplicationUser> signInManager,
            ApplicationDbContext db,
            IEmailSender emailSender,
            IStringLocalizer<SharedResource> s,
            IOptionsMonitor<DataProtectionTokenProviderOptions> tokenOpts,
            IWebHostEnvironment env) // 2025.12.02 Added: 서명 저장용 환경 주입)
        {
            _userManager = userManager;
            _signInManager = signInManager;
            _db = db;
            _emailSender = emailSender;
            _S = s;
            _tokenOpts = tokenOpts;
            _env = env; // 2025.12.02 Added
        }

        public class InputModel
        {
            [EmailAddress(ErrorMessage = "CP_Email_Invalid")]
            [Display(Name = "CP_Email_Label")]
            [Required(ErrorMessage = "CP_Email_Req")]
            public string? Email { get; set; }

            [Display(Name = "_CM_Label_ID")]
            public string? UserName { get; set; }

            [Display(Name = "_CM_Label_Department")]
            [Required(ErrorMessage = "_Alert_Require_ValidDepartment")]
            public int? DepartmentId { get; set; }

            [Display(Name = "_CM_Label_Position")]
            [Required(ErrorMessage = "_Alert_Require_ValidPosition")]
            public int? PositionId { get; set; }

            [Display(Name = "_CM_Label_Name")]
            [Required(ErrorMessage = "_Alert_Require_Name")]
            public string? DisplayName { get; set; }

            [Phone(ErrorMessage = "CP_Phone_Invalid")]
            [Display(Name = "_CM_Label_Phone")]
            public string? PhoneNumber { get; set; }

            [Display(Name = "_CM_Site_Label")]
            [Required(ErrorMessage = "_Alert_Require_ValidSite")]
            public string? CompCd { get; set; }

            [DataType(DataType.Password), Display(Name = "CP_CurrentPwd_Label")]
            public string? CurrentPassword { get; set; }

            [DataType(DataType.Password), Display(Name = "CP_NewPwd_Label")]
            public string? NewPassword { get; set; }

            [DataType(DataType.Password), Display(Name = "CP_NewPwdConfirm_Label")]
            [Compare("NewPassword", ErrorMessage = "CP_NewPwd_NotMatch")]
            public string? ConfirmNewPassword { get; set; }

            // 2025.12.02 Added: 프로필 서명 이미지 업로드 파일
            [Display(Name = "CP_Signature_Label")]
            public IFormFile? SignatureFile { get; set; }
        }

        public class EmailChangeInput
        {
            [Required(ErrorMessage = "_Alert_Require_NewMail")]
            [EmailAddress(ErrorMessage = "CP_Email_Invalid")]
            [Display(Name = "CP_EmailChange_NewEmail_Label")]
            public string? NewEmail { get; set; }

            [Required(ErrorMessage = "CP_EmailChange_CurrentPwd_Req")]
            [DataType(DataType.Password)]
            [Display(Name = "CP_EmailChange_CurrentPwd_Label")]
            public string? CurrentPassword { get; set; }
        }

        [BindProperty] public InputModel Input { get; set; } = new();
        public EmailChangeInput? EmailChange { get; set; } = new();

        public SelectList CompOptions { get; set; } = default!;
        public SelectList DepartmentOptions { get; set; } = default!;
        public SelectList PositionOptions { get; set; } = default!;

        private async Task LoadParentFormAsync(ApplicationUser user)
        {
            var profile = await _db.UserProfiles
                .AsNoTracking()
                .SingleOrDefaultAsync(p => p.UserId == user.Id);

            Input = new InputModel
            {
                Email = user.Email,
                UserName = user.UserName,
                PhoneNumber = user.PhoneNumber,
                DisplayName = profile?.DisplayName,
                DepartmentId = profile?.DepartmentId,
                PositionId = profile?.PositionId,
                CompCd = (profile?.CompCd ?? string.Empty).Trim().ToUpperInvariant()
            };

            await LoadOptionsAsync(); // 2026.05.07 Changed: CompCd 세팅 후 드롭다운 로드

            // 2025.12.02 Added: 사용자 서명 이미지 미리보기 경로 설정
            if (!string.IsNullOrEmpty(profile?.SignatureRelativePath))
            {
                SignatureImageUrl = SignatureRequestBasePath + "/" + profile.SignatureRelativePath;
            }
            else
            {
                SignatureImageUrl = null;
            }
        }

        public async Task<IActionResult> OnGetAsync()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user is null) return RedirectToPage("/Account/Login");

            var profile = await _db.UserProfiles.AsNoTracking()
                .SingleOrDefaultAsync(p => p.UserId == user.Id);

            Input = new InputModel
            {
                Email = user.Email,
                UserName = user.UserName,
                PhoneNumber = user.PhoneNumber,
                DisplayName = profile?.DisplayName,
                DepartmentId = profile?.DepartmentId,
                PositionId = profile?.PositionId,
                CompCd = (profile?.CompCd ?? string.Empty).Trim().ToUpperInvariant()
            };

            await LoadOptionsAsync(); // 2026.05.07 Changed: CompCd 세팅 후 드롭다운 로드

            // 2025.12.02 Added: 첫 진입 시 서명 이미지 미리보기 경로 설정
            if (!string.IsNullOrEmpty(profile?.SignatureRelativePath))
            {
                SignatureImageUrl = SignatureRequestBasePath + "/" + profile.SignatureRelativePath;
            }
            else
            {
                SignatureImageUrl = null;
            }

            EmailChange = new EmailChangeInput();
            return Page();
        }

        // 부모 저장
        public async Task<IActionResult> OnPostSaveAsync()
        {
            // 0) 모달 키 전부 제거
            foreach (var key in ModelState.Keys
                         .Where(k => k.StartsWith("EmailChange.", StringComparison.OrdinalIgnoreCase)
                                  || string.Equals(k, "EmailChange", StringComparison.OrdinalIgnoreCase))
                         .ToList())
            {
                ModelState.Remove(key);
            }

            var user = await _userManager.GetUserAsync(User);
            if (user is null) return RedirectToPage("/Account/Login");

            // 2026.05.07 Changed: 일반 사용자 프로필 저장 시 사업장은 기존 프로필 사업장으로 고정
            var existingProfile = await _db.UserProfiles
                .AsNoTracking()
                .SingleOrDefaultAsync(p => p.UserId == user.Id);

            var fixedCompCd = (existingProfile?.CompCd ?? string.Empty).Trim().ToUpperInvariant();

            if (string.IsNullOrWhiteSpace(fixedCompCd))
            {
                Input.CompCd = fixedCompCd;
                AddErrorOnce($"{nameof(Input)}.{nameof(Input.CompCd)}", _S["_Alert_Require_ValidSite"].Value);
                await LoadOptionsAsync();
                EmailChange = new EmailChangeInput();
                return Page();
            }

            Input.CompCd = fixedCompCd;
            ModelState.Remove($"{nameof(Input)}.{nameof(Input.CompCd)}");

            await LoadOptionsAsync();

            await ValidateMasterSelectionAsync(); // 2026.05.07 Added: 사업장 소속 부서 직급 검증

            // 2025.12.02 Added: 기존 서명 경로 로드하여 오류 시에도 미리보기 유지
            if (!string.IsNullOrEmpty(existingProfile?.SignatureRelativePath))
            {
                SignatureImageUrl = SignatureRequestBasePath + "/" + existingProfile.SignatureRelativePath;
            }

            Input.NewPassword = (Input.NewPassword ?? "").Trim();
            Input.ConfirmNewPassword = (Input.ConfirmNewPassword ?? "").Trim();
            Input.CurrentPassword = (Input.CurrentPassword ?? "").Trim();

            // 1) 비밀번호 변경 의사 판단
            var wantsPwChange = (Input.NewPassword?.Length ?? 0) > 0 ||
                                (Input.ConfirmNewPassword?.Length ?? 0) > 0 ||
                                (Input.CurrentPassword?.Length ?? 0) > 0;

            if (wantsPwChange)
            {
                // 필수 체크
                if (string.IsNullOrWhiteSpace(Input.CurrentPassword))
                    AddErrorOnce($"{nameof(Input)}.{nameof(Input.CurrentPassword)}", _S["CP_CurrentPwd_Req"].Value);

                if (string.IsNullOrWhiteSpace(Input.NewPassword))
                    AddErrorOnce($"{nameof(Input)}.{nameof(Input.NewPassword)}", _S["CP_NewPwd_Req"].Value);

                if (string.IsNullOrWhiteSpace(Input.ConfirmNewPassword))
                    AddErrorOnce($"{nameof(Input)}.{nameof(Input.ConfirmNewPassword)}", _S["CP_NewPwdConfirm_Req"].Value);

                if (!string.IsNullOrEmpty(Input.NewPassword) &&
                    !string.IsNullOrEmpty(Input.ConfirmNewPassword) &&
                    Input.NewPassword != Input.ConfirmNewPassword)
                {
                    AddErrorOnce($"{nameof(Input)}.{nameof(Input.ConfirmNewPassword)}", _S["CP_NewPwd_NotMatch"].Value);
                }

                if (!ModelState.IsValid)
                {
                    EmailChange = new EmailChangeInput();
                    return Page();
                }

                // 현재 비밀번호 확인
                var ok = await _userManager.CheckPasswordAsync(user, Input.CurrentPassword!);
                if (!ok)
                {
                    AddErrorOnce($"{nameof(Input)}.{nameof(Input.CurrentPassword)}", _S["CP_CurrentPwd_Wrong"].Value);
                    EmailChange = new EmailChangeInput();
                    return Page();
                }

                // 실제 변경 (※ 한 번만!)
                var cr = await _userManager.ChangePasswordAsync(user, Input.CurrentPassword!, Input.NewPassword!);
                if (!cr.Succeeded)
                {
                    foreach (var e in cr.Errors)
                        AddErrorOnce($"{nameof(Input)}.{nameof(Input.NewPassword)}", e.Description); // 필드 키 귀속
                    EmailChange = new EmailChangeInput();
                    return Page();
                }
            }

            // 2) 프로필/연락처 저장
            if (!ModelState.IsValid)
            {
                EmailChange = new EmailChangeInput();
                return Page();
            }

            var profile = await _db.UserProfiles.SingleOrDefaultAsync(p => p.UserId == user.Id);
            if (profile is null)
            {
                profile = new UserProfile
                {
                    UserId = user.Id,
                    CompCd = fixedCompCd
                };
                _db.UserProfiles.Add(profile);
            }
            else
            {
                profile.CompCd = fixedCompCd;
            }

            profile.DisplayName = Input.DisplayName?.Trim();
            profile.DepartmentId = Input.DepartmentId;
            profile.PositionId = Input.PositionId;

            // 2025.12.10 Changed: 서명 이미지 파일을 ContentRoot\App_Data\Signatures 아래에 저장
            if (Input.SignatureFile != null && Input.SignatureFile.Length > 0)
            {
                var compCd = (profile.CompCd ?? fixedCompCd).Trim().ToUpperInvariant();
                if (!string.IsNullOrEmpty(compCd))
                {
                    // ★ 변경: WebRootPath → ContentRootPath + App_Data/Signatures
                    var root = Path.Combine(
                        _env.ContentRootPath,
                        SignatureRootRelativeFolder.Replace('/', Path.DirectorySeparatorChar));

                    var compDir = Path.Combine(root, compCd);
                    Directory.CreateDirectory(compDir);

                    const int maxSize = 300;

                    await using (var srcStream = Input.SignatureFile.OpenReadStream())
                    {
                        using var image = Image.Load(srcStream);

                        // 비율 유지 + 최대 300px로 축소(원본이 작으면 그대로)
                        image.Mutate(ctx => ctx.Resize(new ResizeOptions
                        {
                            Mode = ResizeMode.Max,
                            Size = new Size(maxSize, maxSize)
                        }));

                        var fileName = $"{user.Id}.png";
                        var physicalPath = Path.Combine(compDir, fileName);

                        await image.SaveAsync(physicalPath, new PngEncoder());

                        var relativePath = $"{compCd}/{fileName}";
                        profile.SignatureRelativePath = relativePath;
                        SignatureImageUrl = SignatureRequestBasePath + "/" + relativePath;
                    }
                }
            }

            // 이메일/전화는 변경된 경우에만 호출
            if (!string.Equals(user.Email, Input.Email, StringComparison.OrdinalIgnoreCase))
            {
                var er = await _userManager.SetEmailAsync(user, Input.Email);
                if (!er.Succeeded)
                {
                    foreach (var e in er.Errors) AddErrorOnce("Input.Email", e.Description);
                    EmailChange = new EmailChangeInput();
                    return Page();
                }
            }

            if (!string.Equals(user.PhoneNumber, Input.PhoneNumber, StringComparison.Ordinal))
            {
                var pr = await _userManager.SetPhoneNumberAsync(user, Input.PhoneNumber);
                if (!pr.Succeeded)
                {
                    foreach (var e in pr.Errors) AddErrorOnce("Input.PhoneNumber", e.Description);
                    EmailChange = new EmailChangeInput();
                    return Page();
                }
            }

            var ur = await _userManager.UpdateAsync(user);
            if (!ur.Succeeded)
            {
                foreach (var e in ur.Errors) AddErrorOnce("Input.DisplayName", e.Description);
                EmailChange = new EmailChangeInput();
                return Page();
            }

            await _db.SaveChangesAsync();
            await _signInManager.RefreshSignInAsync(user);

            TempData["StatusMessage"] = _S["_CM_Save_Success"].Value;
            return RedirectToPage();
        }

        // 2025.12.02 Added: 서명 이미지 삭제 전용 핸들러
        public async Task<IActionResult> OnPostDeleteSignatureAsync()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user is null) return RedirectToPage("/Account/Login");

            var profile = await _db.UserProfiles
                .SingleOrDefaultAsync(p => p.UserId == user.Id);

            if (profile != null && !string.IsNullOrEmpty(profile.SignatureRelativePath))
            {
                // ★ 변경: WebRootPath → ContentRootPath + App_Data/Signatures
                var root = Path.Combine(
                    _env.ContentRootPath,
                    SignatureRootRelativeFolder.Replace('/', Path.DirectorySeparatorChar));

                var physicalPath = Path.Combine(
                    root,
                    profile.SignatureRelativePath.Replace('/', Path.DirectorySeparatorChar));

                try
                {
                    if (System.IO.File.Exists(physicalPath))
                    {
                        System.IO.File.Delete(physicalPath);
                    }
                }
                catch
                {
                    // 필요 시 로깅 추가
                }

                profile.SignatureRelativePath = null;
                await _db.SaveChangesAsync();
            }

            await LoadParentFormAsync(user);
            SignatureImageUrl = null;

            TempData["StatusMessage"] = _S["CP_Signature_Delete_Done"].Value;
            return RedirectToPage();
        }

        // 이메일 변경(모달)
        public async Task<IActionResult> OnPostSendEmailChangeAsync([FromForm] EmailChangeInput emailChange)
        {
            ModelState.Clear();

            // 모달 필드만 검증
            if (!TryValidateModel(emailChange, "EmailChange"))
            {
                var u = await _userManager.GetUserAsync(User);
                if (u is null) return RedirectToPage("/Account/Login");

                await LoadParentFormAsync(u);
                EmailChange = emailChange;
                ViewData["OpenEmailModal"] = true;
                return Page();
            }
            // 1) 현재 사용자 로드 (이미 있으므로 재사용)
            var user = await _userManager.GetUserAsync(User);
            if (user is null) return RedirectToPage("/Account/Login");

            await LoadOptionsAsync();

            // 2) 비밀번호 확인 (기존 코드 유지)
            var pwOk = await _userManager.CheckPasswordAsync(user, emailChange.CurrentPassword!);
            if (!pwOk)
            {
                await LoadParentFormAsync(user);
                AddErrorOnce("EmailChange.CurrentPassword", _S["CP_EmailChange_CurrentPwd_Wrong"].Value);
                EmailChange = emailChange;
                ViewData["OpenEmailModal"] = true;
                return Page();
            }

            // 3) 동일 이메일인지 (기존 코드 유지)
            if (string.Equals(user.Email, emailChange.NewEmail, StringComparison.OrdinalIgnoreCase))
            {
                await LoadParentFormAsync(user);
                AddErrorOnce("EmailChange.NewEmail", _S["CP_EmailChange_SameAsOld"].Value);
                EmailChange = emailChange;
                ViewData["OpenEmailModal"] = true;
                return Page();
            }

            // 4) 새 이메일이 이미 사용 중인지 확인 (추가)
            var trimmedNew = (emailChange.NewEmail ?? string.Empty).Trim();

            // UserManager는 NormalizedEmail 기준으로 찾으므로 FindByEmailAsync 사용
            var existing = await _userManager.FindByEmailAsync(trimmedNew);
            if (existing != null && existing.Id != user.Id)
            {
                // 다른 사용자가 이미 사용 중
                await LoadParentFormAsync(user); // 부모 폼 값 유지
                AddErrorOnce("EmailChange.NewEmail", _S["CP_EmailChange_InUse"].Value /* 리소스 키 준비 필요 */);
                EmailChange = emailChange;       // 모달 입력 유지
                ViewData["OpenEmailModal"] = true;
                return Page();
            }

            // 5) 이메일 변경 토큰 생성 및 전송

            var token = await _userManager.GenerateChangeEmailTokenAsync(user, emailChange.NewEmail!);
            var code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(token));
            var callbackUrl = Url.Page("/Account/ConfirmEmailChange", null,
                new { userId = user.Id, email = emailChange.NewEmail, code }, Request.Scheme);

            var displayName = (await _db.UserProfiles
                .AsNoTracking()
                .Where(p => p.UserId == user.Id)
                .Select(p => p.DisplayName)
                .FirstOrDefaultAsync()) ?? user.UserName ?? "User";

            var providerName = _userManager.Options.Tokens.ChangeEmailTokenProvider;
            // 이름별 옵션을 조회하되, 실패 시 CurrentValue로 폴백
            var lifeSpan = _tokenOpts.CurrentValue.TokenLifespan;   // ← 이 한 줄만 사용
            var expireMinutes = (int)Math.Round(lifeSpan.TotalMinutes);

            string subject = _S["CP_EmailChange_Subject"].Value;
            string body = _S["CP_EmailChange_BodyHtml",
                HtmlEncoder.Default.Encode(displayName),
                HtmlEncoder.Default.Encode(emailChange.NewEmail!),
                HtmlEncoder.Default.Encode(callbackUrl!),
                expireMinutes].Value;

            await _emailSender.SendEmailAsync(emailChange.NewEmail!, subject, body);

            TempData["StatusMessage"] = _S["CP_EmailChange_Mail_Sent"].Value;
            return RedirectToPage();
        }

        public async Task<IActionResult> OnGetDepartmentsAsync(string compCd)
        {
            var lang = Ui2();
            compCd = (compCd ?? string.Empty).Trim().ToUpperInvariant();

            if (string.IsNullOrWhiteSpace(compCd))
                return new JsonResult(Array.Empty<object>());

            var rows = await _db.DepartmentMasters
                .Where(d => d.IsActive && d.CompCd == compCd)
                .Select(d => new
                {
                    d.Id,
                    d.SortOrder,
                    Fallback = d.Name,
                    LocName = _db.DepartmentMasterLoc
                        .Where(l => l.DepartmentId == d.Id && l.LangCode == lang)
                        .Select(l => l.Name)
                        .FirstOrDefault()
                })
                .AsNoTracking()
                .ToListAsync();

            var items = rows
                .Select(d => new
                {
                    value = d.Id.ToString(),
                    text = string.IsNullOrWhiteSpace(d.LocName) ? d.Fallback : d.LocName,
                    d.SortOrder
                })
                .OrderBy(x => x.SortOrder)
                .ThenBy(x => x.text)
                .ToList();

            return new JsonResult(items);
        }

        public async Task<IActionResult> OnGetPositionsAsync(string compCd)
        {
            var lang = Ui2();
            compCd = (compCd ?? string.Empty).Trim().ToUpperInvariant();

            if (string.IsNullOrWhiteSpace(compCd))
                return new JsonResult(Array.Empty<object>());

            var rows = await (
                from p in _db.PositionMasters
                where p.IsActive && p.CompCd == compCd
                join l in _db.PositionMasterLoc.Where(x => x.LangCode == lang)
                    on p.Id equals l.PositionId into gj
                from l in gj.DefaultIfEmpty()
                orderby p.RankLevel, p.Id
                select new
                {
                    p.Id,
                    p.RankLevel,
                    Fallback = p.Name,
                    LocName = l != null ? l.Name : null,
                    ShortName = l != null ? l.ShortName : null
                }
            ).AsNoTracking().ToListAsync();

            var items = rows
                .Select(p => new
                {
                    value = p.Id.ToString(),
                    text = !string.IsNullOrWhiteSpace(p.ShortName) ? p.ShortName
                         : !string.IsNullOrWhiteSpace(p.LocName) ? p.LocName
                         : p.Fallback,
                    p.RankLevel
                })
                .OrderBy(x => x.RankLevel)
                .ThenBy(x => x.text)
                .ToList();

            return new JsonResult(items);
        }

        private async Task LoadOptionsAsync()
        {
            var lang = Ui2();
            var input = Input ?? new InputModel();

            var compCd = (input.CompCd ?? string.Empty).Trim().ToUpperInvariant();
            var selectedDepartmentId = input.DepartmentId;
            var selectedPositionId = input.PositionId;

            var comps = await _db.Set<CompMaster>()
                .Where(c => c.IsActive)
                .OrderBy(c => c.CompCd)
                .Select(c => new { Value = c.CompCd, Text = c.Name })
                .AsNoTracking()
                .ToListAsync();

            CompOptions = new SelectList(comps, "Value", "Text", compCd);

            if (string.IsNullOrWhiteSpace(compCd))
            {
                DepartmentOptions = new SelectList(new List<SelectListItem>(), "Value", "Text", selectedDepartmentId?.ToString());
                PositionOptions = new SelectList(new List<SelectListItem>(), "Value", "Text", selectedPositionId?.ToString());
                return;
            }

            var deps = await _db.Set<DepartmentMaster>()
                .Where(d => d.IsActive && d.CompCd == compCd)
                .Select(d => new
                {
                    d.Id,
                    d.SortOrder,
                    Fallback = d.Name,
                    LocName = _db.Set<DepartmentMasterLoc>()
                        .Where(l => l.DepartmentId == d.Id && l.LangCode == lang)
                        .Select(l => l.Name)
                        .FirstOrDefault()
                })
                .AsNoTracking()
                .ToListAsync();

            var depItems = deps
                .OrderBy(d => d.SortOrder)
                .ThenBy(d => string.IsNullOrWhiteSpace(d.LocName) ? d.Fallback : d.LocName)
                .Select(d => new SelectListItem
                {
                    Value = d.Id.ToString(),
                    Text = string.IsNullOrWhiteSpace(d.LocName) ? d.Fallback : d.LocName,
                    Selected = selectedDepartmentId is int deptId && d.Id == deptId
                })
                .ToList();

            DepartmentOptions = new SelectList(depItems, "Value", "Text", selectedDepartmentId?.ToString());

            var posRows = await (
                from p in _db.Set<PositionMaster>()
                where p.IsActive && p.CompCd == compCd
                join l in _db.Set<PositionMasterLoc>().Where(x => x.LangCode == lang)
                    on p.Id equals l.PositionId into gj
                from l in gj.DefaultIfEmpty()
                orderby p.RankLevel, p.Id
                select new
                {
                    p.Id,
                    p.RankLevel,
                    Fallback = p.Name,
                    LocName = l != null ? l.Name : null,
                    ShortName = l != null ? l.ShortName : null
                }
            )
            .AsNoTracking()
            .ToListAsync();

            var posItems = posRows
                .OrderBy(p => p.RankLevel)
                .ThenBy(p => !string.IsNullOrWhiteSpace(p.ShortName) ? p.ShortName
                    : !string.IsNullOrWhiteSpace(p.LocName) ? p.LocName
                    : p.Fallback)
                .Select(p => new SelectListItem
                {
                    Value = p.Id.ToString(),
                    Text = !string.IsNullOrWhiteSpace(p.ShortName) ? p.ShortName
                         : !string.IsNullOrWhiteSpace(p.LocName) ? p.LocName
                         : p.Fallback,
                    Selected = selectedPositionId is int posId && p.Id == posId
                })
                .ToList();

            PositionOptions = new SelectList(posItems, "Value", "Text", selectedPositionId?.ToString());
        }

        private async Task ValidateMasterSelectionAsync()
        {
            const string keyCompCd = "Input.CompCd";
            const string keyDepartmentId = "Input.DepartmentId";
            const string keyPositionId = "Input.PositionId";

            var input = Input ?? new InputModel();

            var compCd = (input.CompCd ?? string.Empty).Trim().ToUpperInvariant();
            var departmentId = input.DepartmentId;
            var positionId = input.PositionId;

            var compExists = !string.IsNullOrWhiteSpace(compCd) &&
                await _db.Set<CompMaster>()
                    .AnyAsync(x => x.IsActive && x.CompCd == compCd);

            if (!compExists)
            {
                AddErrorOnce(keyCompCd, _S["_Alert_Require_ValidSite"].Value);
            }

            if (departmentId is int deptId)
            {
                var departmentExists = await _db.Set<DepartmentMaster>()
                    .AnyAsync(x =>
                        x.IsActive &&
                        x.Id == deptId &&
                        x.CompCd == compCd);

                if (!departmentExists)
                {
                    AddErrorOnce(keyDepartmentId, _S["_Alert_Require_ValidDepartment"].Value);
                }
            }

            if (positionId is int posId)
            {
                var positionExists = await _db.Set<PositionMaster>()
                    .AnyAsync(x =>
                        x.IsActive &&
                        x.Id == posId &&
                        x.CompCd == compCd);

                if (!positionExists)
                {
                    AddErrorOnce(keyPositionId, _S["_Alert_Require_ValidPosition"].Value);
                }
            }
        }

        private void AddErrorOnce(string key, string message)
        {
            if (!ModelState.TryGetValue(key, out var entry) || !entry.Errors.Any(e => e.ErrorMessage == message))
                ModelState.AddModelError(key, message);
        }
    }
}