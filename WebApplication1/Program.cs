using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services; // SmtpEmailSender, CustomUserClaimsPrincipalFactory 등
using Fido2NetLib;
using Microsoft.AspNetCore.Localization;
using System.Globalization;
using Microsoft.Extensions.Options;
using WebApplication1;
using Microsoft.AspNetCore.WebUtilities; // QueryHelpers

var builder = WebApplication.CreateBuilder(args);

// -----------------------------
// 1) 외부 서비스 구성
// -----------------------------
builder.Services.Configure<SmtpOptions>(builder.Configuration.GetSection("Smtp"));
builder.Services.AddTransient<IEmailSender, SmtpEmailSender>();
builder.Services.AddScoped<WebAuthnService>();

// -----------------------------
// 2) DbContext
// -----------------------------
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

// -----------------------------
// 3) Identity (딱 한 번만 등록)
// -----------------------------
builder.Services
    .AddIdentity<ApplicationUser, IdentityRole>(options =>
    {
        options.SignIn.RequireConfirmedAccount = false;
        options.Password.RequireNonAlphanumeric = true;
        options.Password.RequireUppercase = true;
        options.Password.RequiredLength = 4;
    })
    .AddEntityFrameworkStores<ApplicationDbContext>()
    .AddDefaultTokenProviders() // UI는 직접 구현 -> AddDefaultUI() 불필요하나, 있어도 무방
    .AddDefaultUI();

// 익명 허용: 필요한 Identity 페이지만 (폴더 전체 허용 제거)
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
    options.Conventions.AllowAnonymousToPage("/Account/AccessDenied"); // 커스텀 접근거부 페이지도 쓰면 함께
});

// 쿠키 설정(단일 블록, 변수명 options 유지)
builder.Services.ConfigureApplicationCookie(options =>
{
    options.LoginPath = "/Account/Login";                // ← 단일 경로
    options.AccessDeniedPath = "/Account/AccessDenied";
    options.ReturnUrlParameter = "returnUrl";

    // 만료/쿠키
    options.SlidingExpiration = true;
    options.ExpireTimeSpan = TimeSpan.FromHours(8);
    options.Cookie.HttpOnly = true;
    options.Cookie.SameSite = SameSiteMode.Lax;
    options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;

    // 어디서든 로그인 리다이렉트 발생 시 /Account/Login 으로
    options.Events.OnRedirectToLogin = context =>
    {
        var returnUrl = context.Request.PathBase + context.Request.Path + context.Request.QueryString;
        var loginUrl = QueryHelpers.AddQueryString(options.LoginPath, options.ReturnUrlParameter, returnUrl);
        context.Response.Redirect(loginUrl);
        return Task.CompletedTask;
    };
});

// 이메일/비번재설정 등 토큰 유효시간
builder.Services.Configure<DataProtectionTokenProviderOptions>(o =>
{
    o.TokenLifespan = TimeSpan.FromMinutes(30);
});

// 커스텀 클레임 팩토리(있을 때만)
builder.Services.AddScoped<IUserClaimsPrincipalFactory<ApplicationUser>, CustomUserClaimsPrincipalFactory>();

// -----------------------------
// 4) Authorization
// -----------------------------
builder.Services.AddAuthorization(options =>
{
    // 로그인 필수(기본 정책)
    options.FallbackPolicy = new AuthorizationPolicyBuilder()
        .RequireAuthenticatedUser()
        .Build();

    // 관리자 정책
    options.AddPolicy("AdminOnly", policy =>
        policy.RequireAssertion(ctx =>
        {
            var v = ctx.User.FindFirst("is_admin")?.Value;
            return v == "1" || v == "2"; // 1: 관리자, 2: 슈퍼관리자
        }));
});



// -----------------------------
// 6) FIDO2
// -----------------------------
builder.Services.AddSingleton(provider =>
{
    var cfg = new Fido2Configuration
    {
        ServerDomain = "localhost",
        ServerName = "Han Young E-Board",
        Origins = new HashSet<string> { "https://localhost:7242" },
        TimestampDriftTolerance = 300_000
    };
    return new Fido2(cfg);
});

// Multi Language
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

    // Cookie > QueryString 순으로 탐지
    options.RequestCultureProviders = new List<IRequestCultureProvider>
    {
        new CookieRequestCultureProvider(),
        new QueryStringRequestCultureProvider()
    };
});

var app = builder.Build();

app.UseRequestLocalization(app.Services.GetRequiredService<IOptions<RequestLocalizationOptions>>().Value);

// -----------------------------
// 7) Pipeline
// -----------------------------
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
    app.UseMigrationsEndPoint();
}
else
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

// 기존 /Identity/Account/Login 접근 차단 → /Account/Login 으로 영구 리다이렉트
app.Use(async (ctx, next) =>
{
    if (ctx.Request.Path.Equals("/Identity/Account/Login", StringComparison.OrdinalIgnoreCase))
    {
        var to = QueryHelpers.AddQueryString("/Account/Login",
                  "returnUrl",
                  ctx.Request.Query["returnUrl"].FirstOrDefault()
                  ?? (ctx.Request.PathBase + ctx.Request.Path + ctx.Request.QueryString));
        ctx.Response.Redirect(to, permanent: true); // 301/308 리다이렉트
        return;
    }
    await next();
});

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

// -----------------------------
// 9) 라우팅
// -----------------------------
app.MapControllerRoute(
    name: "areas",
    pattern: "{area:exists}/{controller=Home}/{action=Index}/{id?}");

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapRazorPages();

app.Run();
