// 2026.03.03 Changed: DocumentTemplatesDX 라우트 유입 및 엔드포인트 등록 여부를 즉시 확인할 수 있도록 ping 및 endpoints 진단 엔드포인트를 추가함
// 2026.02.27 Changed 응답 압축과 정적 파일 캐시 헤더를 추가하여 DX Spreadsheet 초기 로딩을 개선
using DevExpress.AspNetCore;
using Fido2NetLib;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Localization;
using Microsoft.AspNetCore.ResponseCompression;
using Microsoft.AspNetCore.Routing; // EndpointDataSource
using Microsoft.AspNetCore.WebUtilities; // QueryHelpers
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using WebApplication1;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services; // SmtpEmailSender, CustomUserClaimsPrincipalFactory, DocTemplateService


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
builder.Services.AddHostedService<ComposeDxTempCleanupService>();

// -----------------------------
// 10) WebPushNotifier
// -----------------------------
builder.Services.AddScoped<WebPushNotifier>();
builder.Services.AddScoped<WebApplication1.Services.IWebPushNotifier, WebApplication1.Services.WebPushNotifier>();

// -----------------------------
// 11) DevExpress (Build 이전 필수)
// -----------------------------
builder.Services.AddDevExpressControls();

// -----------------------------
// 12) Response Compression (DX Spreadsheet 초기 로딩 개선)
// -----------------------------
builder.Services.AddResponseCompression(options =>
{
    options.EnableForHttps = true;
    options.MimeTypes = ResponseCompressionDefaults.MimeTypes.Concat(new[]
    {
        "application/json",
        "text/plain",
        "application/octet-stream"
    });
});

builder.Services.Configure<BrotliCompressionProviderOptions>(o =>
{
    o.Level = CompressionLevel.Fastest;
});
builder.Services.Configure<GzipCompressionProviderOptions>(o =>
{
    o.Level = CompressionLevel.Fastest;
});

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

app.Use(async (context, next) =>
{
    var path = context.Request.Path.Value ?? "";

    // ✅ 어떤 라우트 형태든 잡히게 완화
    if (path.IndexOf("DxSpreadsheetRequest", StringComparison.OrdinalIgnoreCase) >= 0)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            await next();
        }
        finally
        {
            sw.Stop();
            System.Diagnostics.Debug.WriteLine($"[DxSpreadsheetRequest] {context.Response?.StatusCode} {sw.ElapsedMilliseconds}ms path={path}");
        }
        return;
    }

    await next();
});

// ResponseCompression은 정적 파일보다 위에 두는 편이 효과가 큼
app.UseResponseCompression();

// DX 리소스 캐시 헤더 포함 기본 StaticFiles
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        var p = (ctx.Context.Request.Path.Value ?? "").Replace('\\', '/');

        if (p.StartsWith("/dxresources/", StringComparison.OrdinalIgnoreCase)
            || p.StartsWith("/devex-spreadsheet/", StringComparison.OrdinalIgnoreCase))
        {
            ctx.Context.Response.Headers["Cache-Control"] = "public,max-age=31536000,immutable";
        }
    }
});

// 2026.02.06 Changed: node_modules 위치를 ContentRoot 기준 + 상위 폴더(솔루션 루트)까지 탐색
{
    var contentRoot = app.Environment.ContentRootPath;
    string? nodeModulesAbs = null;

    var here = Path.Combine(contentRoot, "node_modules");
    if (Directory.Exists(here))
    {
        nodeModulesAbs = here;
    }
    else
    {
        var parent = Directory.GetParent(contentRoot)?.FullName;
        if (!string.IsNullOrWhiteSpace(parent))
        {
            var up1 = Path.Combine(parent, "node_modules");
            if (Directory.Exists(up1))
            {
                nodeModulesAbs = up1;
            }
        }
    }

    if (!string.IsNullOrWhiteSpace(nodeModulesAbs))
    {
        app.UseStaticFiles(new StaticFileOptions
        {
            FileProvider = new PhysicalFileProvider(nodeModulesAbs),
            RequestPath = "/node_modules"
        });
    }
}

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

// Development에서만 DXR.axd 예외 진단
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

// ✅ Auth는 엔드포인트 매핑 전에
app.UseAuthentication();
app.UseAuthorization();

//  DevExpress는 Routing 이후에만(Spreadsheet 25.2 기준)
//  MapDevExpressControls() 호출은 제거(현재 참조 패키지에 메서드가 없어서 CS1061 발생)
app.UseDevExpressControls();

// ✅ (추가) 앱 유입/엔드포인트 등록 진단용
app.MapGet("/DocumentTemplatesDX/ping", () => Results.Ok(new { ok = true, t = DateTimeOffset.Now }))
   .AllowAnonymous();

app.MapGet("/__diag/endpoints", (EndpointDataSource ds, string? contains) =>
{
    contains = (contains ?? "").Trim();

    var list = ds.Endpoints
        .Select(e => e.DisplayName ?? "")
        .Where(s => !string.IsNullOrWhiteSpace(s))
        .Where(s => string.IsNullOrEmpty(contains) || s.Contains(contains, StringComparison.OrdinalIgnoreCase))
        .OrderBy(s => s)
        .Take(500)
        .ToArray();

    return Results.Ok(new { count = list.Length, contains, list });
}).AllowAnonymous();

app.MapControllers();

//  MVC 라우트(기존 2개 유지)
app.MapControllerRoute(
    name: "areas",
    pattern: "{area:exists}/{controller=Home}/{action=Index}/{id?}");

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapRazorPages();

app.MapGet("/devex/ping", () => Results.Ok(new { ok = true, devexpress = "enabled" }));

app.Run();