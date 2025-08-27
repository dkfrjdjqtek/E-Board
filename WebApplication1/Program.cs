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
    .AddDefaultTokenProviders() // UI는 직접 구현 -> AddDefaultUI() 불필요
    .AddDefaultUI();

// 이메일/비번재설정 등 토큰 유효시간
builder.Services.Configure<DataProtectionTokenProviderOptions>(o =>
{
    o.TokenLifespan = TimeSpan.FromMinutes(30);
});

// 커스텀 클레임 팩토리(있을 때만)
builder.Services.AddScoped<IUserClaimsPrincipalFactory<ApplicationUser>, CustomUserClaimsPrincipalFactory>();

// -----------------------------
// 4) 쿠키/락아웃
// -----------------------------
builder.Services.ConfigureApplicationCookie(o =>
{
    o.LoginPath = "/Account/Login";
    o.AccessDeniedPath = "/Account/AccessDenied";
    o.ReturnUrlParameter = "returnUrl";
    o.SlidingExpiration = true;
    o.ExpireTimeSpan = TimeSpan.FromHours(8);
    o.Cookie.HttpOnly = true;
    o.Cookie.SameSite = SameSiteMode.Lax;
    o.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
});
builder.Services.Configure<IdentityOptions>(o =>
{
    o.Lockout.MaxFailedAccessAttempts = 5;
    o.Lockout.DefaultLockoutTimeSpan = TimeSpan.FromMinutes(10);
});

// -----------------------------
// 5) MVC/Razor + Authorization (하나로 합침)
// -----------------------------
//builder.Services.AddControllersWithViews();
//builder.Services.AddRazorPages();

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

builder.Services.AddRazorPages(options =>
{
    // 꼭 필요한 페이지만 열어둡니다
    options.Conventions.AllowAnonymousToAreaFolder("Identity", "/Account");
    //options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/Login");
    ////options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/Register");
    //options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ForgotPassword");
    //options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ForgotPasswordConfirmation");
    //options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ResetPassword");
    //options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ResetPasswordConfirmation");
    //options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ConfirmEmail");
    //options.Conventions.AllowAnonymousToAreaPage("Identity", "/Account/ConfirmEmailChange");
    
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

//Multi Language
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

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

// -----------------------------
// 8) 데이터 시드 (한 번만 실행)
// -----------------------------
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
    var um = scope.ServiceProvider.GetRequiredService<UserManager<ApplicationUser>>();

    // 연결된 DB 출력(확인용)
    var cnn = db.Database.GetDbConnection();
    app.Logger.LogWarning("Seeding to -> DataSource={DataSource}, Database={Database}",
        cnn.DataSource, cnn.Database);

    await SeedAsync(db, um);
}

static async Task SeedAsync(ApplicationDbContext db, UserManager<ApplicationUser> um)
{
    const string comp = "0001";
    const string adminEmail = "admin@local";
    const string adminPw = "Admin!2345";

    // 없으면 생성(중복 무시)
    await db.Database.ExecuteSqlInterpolatedAsync($@"
IF NOT EXISTS (SELECT 1 FROM dbo.PositionMasters WHERE CompCd = {comp} AND Code = N'E11')
BEGIN
  INSERT dbo.PositionMasters(CompCd, Code, Name, IsActive, RankLevel, IsApprover, SortOrder)
  VALUES ({comp}, N'E11', N'상무', 1, 220, 1, 220);
END");

    await db.Database.ExecuteSqlInterpolatedAsync($@"
IF NOT EXISTS (SELECT 1 FROM dbo.DepartmentMasters WHERE CompCd = {comp} AND Code = N'D006')
BEGIN
  INSERT dbo.DepartmentMasters(CompCd, Code, Name, IsActive, SortOrder)
  VALUES ({comp}, N'D006', N'IT', 1, 60);
END");

    var posId = await db.PositionMasters
        .Where(x => x.CompCd == comp && x.Code == "E11")
        .Select(x => (int?)x.Id)
        .FirstOrDefaultAsync();

    var deptId = await db.DepartmentMasters
        .Where(x => x.CompCd == comp && x.Code == "D006")
        .Select(x => (int?)x.Id)
        .FirstOrDefaultAsync();

    if (posId is null || deptId is null)
        throw new InvalidOperationException("Seed 실패: E11 또는 D006을 다시 확인하세요.");

    var admin = await um.FindByEmailAsync(adminEmail);
    if (admin is null)
    {
        admin = new ApplicationUser { UserName = adminEmail, Email = adminEmail, EmailConfirmed = true };
        await um.CreateAsync(admin, adminPw);
    }

    if (!await db.UserProfiles.AnyAsync(p => p.UserId == admin.Id))
    {
        db.UserProfiles.Add(new UserProfile
        {
            UserId = admin.Id,
            CompCd = comp,
            DisplayName = "관리자",
            DepartmentId = deptId!.Value,
            PositionId = posId!.Value
        });
        await db.SaveChangesAsync();
    }
}

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
