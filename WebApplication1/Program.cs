using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Data;
using WebApplication1.Services; // ���ӽ����̽� ���߱�
using Fido2NetLib;
using Fido2NetLib.Objects;
using Microsoft.Extensions.Options;
using WebApplication1.Models;


var builder = WebApplication.CreateBuilder(args);

builder.Services.Configure<SmtpOptions>(builder.Configuration.GetSection("Smtp"));
builder.Services.AddTransient<IEmailSender, SmtpEmailSender>();
builder.Services.AddScoped<WebAuthnService>();

//// 1) DbContext
//builder.Services.AddDbContext<ApplicationDbContext>(options =>
//    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

//// 2) Identity (MVC ��Ʈ�ѷ� ���)
///* 
//builder.Services
//    .AddIdentity<IdentityUser, IdentityRole>(options => {
//        // ��й�ȣ ��å �� �ɼ�
//        options.SignIn.RequireConfirmedAccount = false;
//        options.Password.RequireNonAlphanumeric = true;
//        options.Password.RequireUppercase = true;
//        options.Password.RequiredLength = 4;
//    })
//    .AddEntityFrameworkStores<ApplicationDbContext>()
//    .AddDefaultTokenProviders()
//    .AddDefaultUI();


//builder.Services
//    .AddDefaultIdentity<ApplicationUser>(options =>
//    {
//        options.SignIn.RequireConfirmedAccount = false;
//    })
//    .AddEntityFrameworkStores<ApplicationDbContext>();
//*/
//builder.Services
//    .AddDefaultIdentity<ApplicationUser>(o => o.SignIn.RequireConfirmedAccount = false)
//    .AddEntityFrameworkStores<ApplicationDbContext>();
//// 3) Cookie ���� ��� ����
//builder.Services.ConfigureApplicationCookie(options => {
//    options.LoginPath = "/Account/Login";
//    options.AccessDeniedPath = "/Account/Login";
//});

//// 4) MVC + ��å
//builder.Services.AddControllersWithViews();
//builder.Services.AddAuthorization(options => {
//    options.FallbackPolicy = new AuthorizationPolicyBuilder()
//        .RequireAuthenticatedUser()
//        .Build();
//});

// 1) DbContext
builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));


// 2) Identity (ApplicationUser ����, �ʿ��ϸ� Role ����)
builder.Services.AddIdentity<ApplicationUser, IdentityRole>(options =>
{
    options.SignIn.RequireConfirmedAccount = false;
    options.Password.RequireNonAlphanumeric = true;
    options.Password.RequireUppercase = true;
    options.Password.RequiredLength = 4;
})
.AddEntityFrameworkStores<ApplicationDbContext>()
.AddDefaultTokenProviders();   // (UI�� ���� ���� View ����̹Ƿ� AddDefaultUI()�� ���ʿ�)

builder.Services.AddScoped<IUserClaimsPrincipalFactory<ApplicationUser>, CustomUserClaimsPrincipalFactory>();

builder.Services.AddAuthorization(options =>
{
    options.AddPolicy("AdminOnly", policy =>
        policy.RequireAssertion(ctx =>
        {
            var v = ctx.User.FindFirst("is_admin")?.Value;
            return v == "1" || v == "2"; // 1:������, 2:����
        }));
});

// 3) ��Ű ���
builder.Services.ConfigureApplicationCookie(options =>
{
    options.LoginPath = "/Account/Login";
    options.AccessDeniedPath = "/Account/AccessDenied";
    options.ReturnUrlParameter = "returnUrl";
    // ��Ű/���ƿ�/�����̵� ���� (���� ���� ���� ���� ���� �ʿ�)
    options.SlidingExpiration = true;
    options.ExpireTimeSpan = TimeSpan.FromHours(8);
    options.Cookie.HttpOnly = true;
    options.Cookie.SameSite = SameSiteMode.Lax;
    options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
});

builder.Services.Configure<IdentityOptions>(o =>
{
    o.Lockout.MaxFailedAccessAttempts = 5;
    o.Lockout.DefaultLockoutTimeSpan = TimeSpan.FromMinutes(10);
});// ��Ű/���ƿ�/�����̵� ���� (���� ���� ���� ���� ���� �ʿ�)

// MVC
builder.Services.AddControllersWithViews();
builder.Services.AddAuthorization(options =>
{
    options.FallbackPolicy = new AuthorizationPolicyBuilder()
        .RequireAuthenticatedUser()
        .Build();
});
// 5) Fido2 ����
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



builder.Services.AddRazorPages();

var app = builder.Build();

// pipeline
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
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
    var um = scope.ServiceProvider.GetRequiredService<UserManager<ApplicationUser>>();

    // �� ���� ����� ����/DB Ȯ��
    var cnn = db.Database.GetDbConnection();
    app.Logger.LogWarning("Seeding to -> DataSource={DataSource}, Database={Database}",
        cnn.DataSource, cnn.Database);

    await SeedAsync(db, um);
}
async Task SeedAsync(ApplicationDbContext db, UserManager<ApplicationUser> um)
{
    // �� ���Ⱑ ������ CS0103 ������ �� �̴ϴ�.
    const string comp = "0001";
    const string adminEmail = "admin@local";
    const string adminPw = "Admin!2345";
    // 1) Ȥ�ö� ������ T-SQL�� ����(�ߺ� ����)
    await db.Database.ExecuteSqlInterpolatedAsync($@"
IF NOT EXISTS (SELECT 1 FROM dbo.PositionMasters WHERE CompCd = {comp} AND Code = N'E11')
BEGIN
  INSERT dbo.PositionMasters(CompCd, Code, Name, IsActive, RankLevel, IsApprover, SortOrder)
  VALUES ({comp}, N'E11', N'��', 1, 220, 1, 220);
END");

    await db.Database.ExecuteSqlInterpolatedAsync($@"
IF NOT EXISTS (SELECT 1 FROM dbo.DepartmentMasters WHERE CompCd = {comp} AND Code = N'D006')
BEGIN
  INSERT dbo.DepartmentMasters(CompCd, Code, Name, IsActive, SortOrder)
  VALUES ({comp}, N'D006', N'IT', 1, 60);
END");

    // 2) �ٽ� �� ���� ��ȸ�ؼ� id Ȯ��
    var posId = await db.PositionMasters
        .Where(x => x.CompCd == comp && x.Code == "E11")
        .Select(x => (int?)x.Id)
        .FirstOrDefaultAsync();

    var deptId = await db.DepartmentMasters
        .Where(x => x.CompCd == comp && x.Code == "D006")
        .Select(x => (int?)x.Id)
        .FirstOrDefaultAsync();

    Console.WriteLine($"[Seed] E11 => {(posId is null ? "NULL" : posId.ToString())}");
    Console.WriteLine($"[Seed] D006 => {(deptId is null ? "NULL" : deptId.ToString())}");

    // 3) ������ null�̸� �ٷ� ����׿� ������ ��� ������ Ȯ��
    if (posId is null || deptId is null)
    {
        var pmDump = await db.PositionMasters
            .Where(x => x.CompCd == comp)
            .Select(x => new { x.Id, x.CompCd, x.Code, x.Name })
            .ToListAsync();
        var dmDump = await db.DepartmentMasters
            .Where(x => x.CompCd == comp)
            .Select(x => new { x.Id, x.CompCd, x.Code, x.Name })
            .ToListAsync();

        Console.WriteLine("[Seed][PM dump] " + string.Join(", ", pmDump.Select(r => $"{r.Id}:{r.CompCd}/{r.Code}/{r.Name}")));
        Console.WriteLine("[Seed][DM dump] " + string.Join(", ", dmDump.Select(r => $"{r.Id}:{r.CompCd}/{r.Code}/{r.Name}")));

        throw new InvalidOperationException("Seed ����: E11 �Ǵ� D006�� �ٽ� Ȯ���ϼ���(�� ���� ����).");
    }

    // --- ������ ����/������ ---
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
            DisplayName = "������",
            DepartmentId = deptId!.Value,
            PositionId = posId!.Value
        });
        await db.SaveChangesAsync();
    }
}


    using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
    var um = scope.ServiceProvider.GetRequiredService<UserManager<ApplicationUser>>();
    await SeedAsync(db, um);
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

// MVC ����ø�
app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");
app.MapRazorPages();
app.Run();
