// 2025.11.10 Changed: Kestrel 바인딩은 launchSettings.json에 위임, Development 환경에서 HTTPS 리다이렉트 해제
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services; // SmtpEmailSender, CustomUserClaimsPrincipalFactory, DocTemplateService
using Fido2NetLib;
using Microsoft.AspNetCore.Localization;
using System.Globalization;
using Microsoft.Extensions.Options;
using WebApplication1;
using Microsoft.AspNetCore.WebUtilities; // QueryHelpers
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);

// -----------------------------
// 1) 외부 서비스 구성
// -----------------------------
builder.Services.Configure<SmtpOptions>(builder.Configuration.GetSection("Smtp"));
builder.Services.AddTransient<IEmailSender, SmtpEmailSender>();
builder.Services.AddScoped<WebAuthnService>();
builder.Services.AddScoped<IAuditLogger, AuditLoggerSql>(); // Build() 이전

// -----------------------------
// 2) DbContext
// -----------------------------
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// -----------------------------
// 3) Identity
// -----------------------------
builder.Services
    .AddIdentity<ApplicationUser, IdentityRole>(options =>
    {
        options.SignIn.RequireConfirmedAccount = false;
        options.Password.RequireNonAlphanumeric = false;
        options.Password.RequireUppercase = false;
        options.Password.RequireLowercase = false;         // 소문자 필수
        options.Password.RequiredLength = 4;
        //options.Password.RequiredUniqueChars = 3;         
    })
    .AddEntityFrameworkStores<ApplicationDbContext>()
    .AddDefaultTokenProviders()
    .AddDefaultUI();

// 익명 허용: 필요한 Identity 페이지만
builder.Services.AddRazorPages(options =>
{
    options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/Login");
    options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ForgotPassword");
    options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ForgotPasswordConfirmation");
    options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ResetPassword");
    options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ResetPasswordConfirmation");
    options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ConfirmEmail");
    options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ConfirmEmailChange");
    options.Conventions.AllowAnonymousToPage("/Account/Login");
    options.Conventions.AllowAnonymousToPage("/Account/AccessDenied");
    options.Conventions.AddPageRoute("/Identity/Admin/AdminUserProfile", "/AdminUserProfile");
});

// 쿠키 설정
builder.Services.ConfigureApplicationCookie(options =>
{
    options.LoginPath = "/Account/Login";
    options.AccessDeniedPath = "/Account/AccessDenied";
    options.ReturnUrlParameter = "returnUrl";

    options.SlidingExpiration = true;
    options.ExpireTimeSpan = TimeSpan.FromHours(8);
    options.Cookie.HttpOnly = true;
    options.Cookie.SameSite = SameSiteMode.Lax;
    options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;

    options.Events.OnRedirectToLogin = context =>
    {
        var returnUrl = context.Request.PathBase + context.Request.Path + context.Request.QueryString;
        var loginUrl = QueryHelpers.AddQueryString(options.LoginPath, options.ReturnUrlParameter, returnUrl);
        context.Response.Redirect(loginUrl);
        return Task.CompletedTask;
    };
});

// 이메일/비번 재설정 등 토큰 유효시간
builder.Services.Configure<DataProtectionTokenProviderOptions>(o =>
{
    o.TokenLifespan = TimeSpan.FromMinutes(30);
});

// 2025.09.26 Added: 초대 완료 전 로그인 방지
builder.Services.Configure<IdentityOptions>(o =>
{
    o.SignIn.RequireConfirmedEmail = true;
});

// 커스텀 클레임 팩토리
builder.Services.AddScoped<IUserClaimsPrincipalFactory<ApplicationUser>, CustomUserClaimsPrincipalFactory>();

// -----------------------------
// 4) Authorization
// -----------------------------
builder.Services.AddAuthorization(options =>
{
    options.FallbackPolicy = new AuthorizationPolicyBuilder()
        .RequireAuthenticatedUser()
        .Build();

    options.AddPolicy("AdminOnly", policy =>
        policy.RequireAssertion(ctx =>
        {
            var v = ctx.User.FindFirst("is_admin")?.Value;
            return v == "1" || v == "2";
        }));
});

// -----------------------------
// 5) 업로드(요청 본문) 한도 — 50MB
// -----------------------------
builder.Services.Configure<FormOptions>(o =>
{
    o.MultipartBodyLengthLimit = 50L * 1024 * 1024; // 50MB
});
builder.Services.Configure<IISServerOptions>(o =>
{
    o.MaxRequestBodySize = 50L * 1024 * 1024; // IIS(in-process)
});
builder.WebHost.ConfigureKestrel(o =>
{
    o.Limits.MaxRequestBodySize = 50L * 1024 * 1024; // Kestrel
});

// -----------------------------
// 6) FIDO2
// -----------------------------
builder.Services.AddSingleton(provider =>
{
    var cfg = new Fido2Configuration
    {
        //ServerDomain = "localhost",
        ServerDomain = "localhost",
        ServerName = "Han Young E-Board",
        Origins = new HashSet<string> { "https://localhost:7242" },
        TimestampDriftTolerance = 300_000
    };
    return new Fido2(cfg);
});

// -----------------------------
// 7) 다국어
// -----------------------------
builder.Services.AddLocalization(options => options.ResourcesPath = "Resources");

builder.Services
    .AddControllersWithViews()
    .AddViewLocalization()
    .AddDataAnnotationsLocalization(options =>
    {
        options.DataAnnotationLocalizerProvider = (type, factory) =>
            factory.Create(typeof(SharedResource));
    });

var cultures = new[] { "ko-KR", "en-US", "vi-VN", "id-ID", "zh-CN" }
    .Select(c => new CultureInfo(c))
    .ToList();

builder.Services.Configure<RequestLocalizationOptions>(options =>
{
    options.DefaultRequestCulture = new RequestCulture("ko-KR");
    options.SupportedCultures = cultures;
    options.SupportedUICultures = cultures;

    options.RequestCultureProviders = new List<IRequestCultureProvider>
    {
        new CookieRequestCultureProvider(),
        new QueryStringRequestCultureProvider()
    };
});

// -----------------------------
// 8) Anti-Forgery
// -----------------------------
builder.Services.AddAntiforgery(options =>
{
    options.HeaderName = "RequestVerificationToken";
});

// -----------------------------
// 9) DocTemplateService
// -----------------------------
builder.Services.AddScoped<IDocTemplateService, DocTemplateService>();

// ===== Build =====
var app = builder.Build();

app.UseRequestLocalization(app.Services.GetRequiredService<IOptions<RequestLocalizationOptions>>().Value);

// -----------------------------
// 10) Pipeline
// -----------------------------
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
    app.UseMigrationsEndPoint();
    // 개발 환경: HTTP 5000 사용, HTTPS 리다이렉트 없음
}
else
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
    //app.UseHttpsRedirection(); // 운영에서만 HTTPS 강제
}

// 10-1) wwwroot 정적 파일 제공 (★ 추가)
app.UseStaticFiles();   // /css, /js, /lib, /images 등 wwwroot 하위 전체

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(
        Path.Combine(app.Environment.ContentRootPath, "App_Data", "Signatures")),
    RequestPath = "/images/signatures"
});

// /Identity/Account/Login -> /Account/Login 영구 리다이렉트
app.Use(async (ctx, next) =>
{
    if (ctx.Request.Path.Equals("/Identity/Account/Login", StringComparison.OrdinalIgnoreCase))
    {
        var to = QueryHelpers.AddQueryString(
            "/Account/Login",
            "returnUrl",
            ctx.Request.Query["returnUrl"].FirstOrDefault()
            ?? (ctx.Request.PathBase + ctx.Request.Path + ctx.Request.QueryString)
        );
        ctx.Response.Redirect(to, permanent: true);
        return;
    }
    await next();
});

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

// -----------------------------
// 11) 라우팅
// -----------------------------
app.MapControllerRoute(
    name: "areas",
    pattern: "{area:exists}/{controller=Home}/{action=Index}/{id?}");

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapRazorPages();

app.Run();