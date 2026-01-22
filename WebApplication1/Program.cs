// 2026.01.21 Changed: DXR.axd 500 원인 파악을 위해 Development에서만 DXR 요청 예외를 본문으로 노출하는 진단 미들웨어를 추가하고, DevExpress 리소스 처리를 위해 UseDevExpressControls 위치를 UseRouting 뒤로 정리
using DevExpress.AspNetCore;
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
using System.Text.Json;

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
        options.Password.RequireLowercase = false;
        options.Password.RequiredLength = 4;
    })
    .AddEntityFrameworkStores<ApplicationDbContext>()
    .AddDefaultTokenProviders()
    .AddDefaultUI();

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

builder.Services.Configure<DataProtectionTokenProviderOptions>(o =>
{
    o.TokenLifespan = TimeSpan.FromMinutes(30);
});

builder.Services.Configure<IdentityOptions>(o =>
{
    o.SignIn.RequireConfirmedEmail = true;
});

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
// 5) 업로드 한도 50MB
// -----------------------------
builder.Services.Configure<FormOptions>(o =>
{
    o.MultipartBodyLengthLimit = 50L * 1024 * 1024;
});
builder.Services.Configure<IISServerOptions>(o =>
{
    o.MaxRequestBodySize = 50L * 1024 * 1024;
});
builder.WebHost.ConfigureKestrel(o =>
{
    o.Limits.MaxRequestBodySize = 50L * 1024 * 1024;
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
    })
    .AddJsonOptions(o =>
    {
        o.JsonSerializerOptions.PropertyNameCaseInsensitive = true;
        o.JsonSerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
        o.JsonSerializerOptions.DictionaryKeyPolicy = JsonNamingPolicy.CamelCase;
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

// -----------------------------
// 10) WebPushNotifier
// -----------------------------
builder.Services.AddScoped<WebPushNotifier>();
builder.Services.AddScoped<WebApplication1.Services.IWebPushNotifier, WebApplication1.Services.WebPushNotifier>();

// -----------------------------
// 11) DevExpress 컨트롤 등록
// -----------------------------
builder.Services.AddDevExpressControls();

var app = builder.Build();

app.UseRequestLocalization(app.Services.GetRequiredService<IOptions<RequestLocalizationOptions>>().Value);

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

app.UseStaticFiles();

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(
        Path.Combine(app.Environment.ContentRootPath, "App_Data", "Signatures")),
    RequestPath = "/images/signatures"
});

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

// ----------------------------------------
// DevExpress 리소스(DXR.axd) 진단: Development에서만 예외 본문 노출
// - 지금은 Response 탭이 비어있어서 원인 추적이 막히므로, 여기서 확정 원인을 뽑습니다.
// ----------------------------------------
if (app.Environment.IsDevelopment())
{
    app.Use(async (ctx, next) =>
    {
        if (ctx.Request.Path.StartsWithSegments("/DXR.axd", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                await next();
            }
            catch (Exception ex)
            {
                ctx.Response.Clear();
                ctx.Response.StatusCode = 500;
                ctx.Response.ContentType = "text/plain; charset=utf-8";
                await ctx.Response.WriteAsync("DXR.axd failed\n");
                await ctx.Response.WriteAsync(ex.GetType().FullName + "\n");
                await ctx.Response.WriteAsync(ex.Message + "\n\n");
                await ctx.Response.WriteAsync(ex.ToString());
                return;
            }
        }
        else
        {
            await next();
        }
    });
}

// DevExpress 미들웨어는 Routing 이후에 두고, DXR.axd / 내부 콜백을 여기서 처리하도록 함
app.UseDevExpressControls();

app.UseAuthentication();
app.UseAuthorization();

app.MapControllerRoute(
    name: "areas",
    pattern: "{area:exists}/{controller=Home}/{action=Index}/{id?}");

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapRazorPages();

// MapDevExpressControls 는 22.2.12에서 확장 메서드가 없는 환경이 있으므로 유지하지 않음
//app.MapDevExpressControls();

app.MapGet("/devex/ping", () => Results.Ok(new { ok = true, devexpress = "enabled" }));

app.Run();
