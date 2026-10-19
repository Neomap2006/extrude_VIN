using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using VelastoProductionSystem.Data;
using VelastoProductionSystem.Helpers;
using VelastoProductionSystem.WebSockets;
using VelastoProductionSystem.Hubs;
using VelastoProductionSystem.Services;
using VelastoProductionSystem.Middleware;

// Set EPPlus License globally for version 8+
OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews(options =>
{
    options.Filters.Add<VelastoProductionSystem.Filters.SessionCheckFilter>();
});
builder.Services.AddSignalR();
builder.Services.AddHttpContextAccessor();
builder.Services.AddScoped<IApprovalService, ApprovalService>();

// Configure Entity Framework Core with SQL Server or SQLite
var defaultConnection = builder.Configuration.GetConnectionString("DefaultConnection");
builder.Services.AddDbContext<ApplicationDbContext>(options =>
{
    if (defaultConnection != null && (defaultConnection.Contains(".db") || defaultConnection.Contains("Data Source")))
        options.UseSqlite(defaultConnection);
    else
        options.UseSqlServer(defaultConnection);
});

// Configure second DB context for ELWP_PRD (read-only, for planning sync)
var elwpConnection = builder.Configuration.GetConnectionString("ElwpConnection");
builder.Services.AddDbContext<ElwpDbContext>(options =>
{
    if (elwpConnection != null && (elwpConnection.Contains(".db") || elwpConnection.Contains("Data Source")))
        options.UseSqlite(elwpConnection);
    else
        options.UseSqlServer(elwpConnection);
});

// Add session support for production tracking
builder.Services.AddSession(options =>
{
    options.IdleTimeout = TimeSpan.FromHours(8);
    options.Cookie.HttpOnly = true;
    options.Cookie.IsEssential = true;
});

var app = builder.Build();

// Ensure databases are created and seed initial data
using (var scope = app.Services.CreateScope())
{
    var services = scope.ServiceProvider;
    try
    {
        var context = services.GetRequiredService<ApplicationDbContext>();
        var elwpContext = services.GetRequiredService<ElwpDbContext>();
        
        // Ensure schemas exist if using SQLite
        if (context.Database.ProviderName?.Contains("Sqlite") == true)
        {
            context.Database.EnsureCreated();
        }
        
        if (elwpContext.Database.ProviderName?.Contains("Sqlite") == true)
        {
            elwpContext.Database.EnsureCreated();
        }

        EnsureApprovalWorkflowObjects(context);
        EnsureSpsDocumentActivationColumn(context);
        EnsureSpsItemListLengthColumns(context);
        EnsureProductionReportExtendedColumns(context);

        DataSeeder.SeedData(context, elwpContext);
        await ShiftHelper.EnsureCompanyShiftsAsync(context);
    }
    catch (Exception ex)
    {
        var logger = services.GetRequiredService<ILogger<Program>>();
        logger.LogError(ex, "An error occurred while initializing the database.");
    }
}

static void EnsureApprovalWorkflowObjects(ApplicationDbContext context)
{
    var provider = context.Database.ProviderName ?? string.Empty;

    if (provider.Contains("SqlServer", StringComparison.OrdinalIgnoreCase))
    {
        context.Database.ExecuteSqlRaw(@"
IF OBJECT_ID(N'[dbo].[ApprovalRequests]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ApprovalRequests] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [RequestCode] NVARCHAR(40) NOT NULL,
        [ActionType] INT NOT NULL,
        [TargetKey] NVARCHAR(256) NOT NULL,
        [RequestComment] NVARCHAR(1024) NULL,
        [ReturnUrl] NVARCHAR(256) NULL,
        [PayloadJson] NVARCHAR(MAX) NULL,
        [RequesterUserName] NVARCHAR(150) NOT NULL,
        [RequesterRole] NVARCHAR(50) NOT NULL,
        [ApproverUserName] NVARCHAR(150) NULL,
        [ApproverRole] NVARCHAR(50) NULL,
        [ApproverComment] NVARCHAR(1024) NULL,
        [ApprovalToken] NVARCHAR(80) NULL,
        [TokenExpiresAt] DATETIME2 NULL,
        [Status] INT NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL,
        [UpdatedAt] DATETIME2 NULL,
        [ReviewedAt] DATETIME2 NULL,
        [ConsumedAt] DATETIME2 NULL
    );
END;

IF COL_LENGTH(N'[dbo].[ApprovalRequests]', N'PayloadJson') IS NULL
BEGIN
    ALTER TABLE [dbo].[ApprovalRequests]
    ADD [PayloadJson] NVARCHAR(MAX) NULL;
END;

IF OBJECT_ID(N'[dbo].[ApprovalRequestLogs]', N'U') IS NULL
BEGIN
    CREATE TABLE [dbo].[ApprovalRequestLogs] (
        [Id] INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [ApprovalRequestId] INT NOT NULL,
        [FromStatus] INT NULL,
        [ToStatus] INT NOT NULL,
        [Comment] NVARCHAR(1024) NULL,
        [ActorUserName] NVARCHAR(150) NOT NULL,
        [ActorRole] NVARCHAR(50) NOT NULL,
        [CreatedAt] DATETIME2 NOT NULL,
        CONSTRAINT [FK_ApprovalRequestLogs_ApprovalRequests_ApprovalRequestId]
            FOREIGN KEY ([ApprovalRequestId]) REFERENCES [dbo].[ApprovalRequests]([Id]) ON DELETE CASCADE
    );
END;

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ApprovalRequests_RequestCode' AND object_id = OBJECT_ID(N'[dbo].[ApprovalRequests]'))
    CREATE UNIQUE INDEX [IX_ApprovalRequests_RequestCode] ON [dbo].[ApprovalRequests]([RequestCode]);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ApprovalRequests_Status_CreatedAt' AND object_id = OBJECT_ID(N'[dbo].[ApprovalRequests]'))
    CREATE INDEX [IX_ApprovalRequests_Status_CreatedAt] ON [dbo].[ApprovalRequests]([Status], [CreatedAt]);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ApprovalRequests_RequesterUserName_Status' AND object_id = OBJECT_ID(N'[dbo].[ApprovalRequests]'))
    CREATE INDEX [IX_ApprovalRequests_RequesterUserName_Status] ON [dbo].[ApprovalRequests]([RequesterUserName], [Status]);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ApprovalRequests_ActionType_TargetKey_RequesterUserName_Status' AND object_id = OBJECT_ID(N'[dbo].[ApprovalRequests]'))
    CREATE INDEX [IX_ApprovalRequests_ActionType_TargetKey_RequesterUserName_Status]
    ON [dbo].[ApprovalRequests]([ActionType], [TargetKey], [RequesterUserName], [Status]);

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_ApprovalRequestLogs_ApprovalRequestId' AND object_id = OBJECT_ID(N'[dbo].[ApprovalRequestLogs]'))
    CREATE INDEX [IX_ApprovalRequestLogs_ApprovalRequestId] ON [dbo].[ApprovalRequestLogs]([ApprovalRequestId]);

IF OBJECT_ID(N'[dbo].[__EFMigrationsHistory]', N'U') IS NOT NULL
   AND NOT EXISTS (SELECT 1 FROM [dbo].[__EFMigrationsHistory] WHERE [MigrationId] = N'20260528072311_AddSpsApprovalWorkflow')
BEGIN
    INSERT INTO [dbo].[__EFMigrationsHistory] ([MigrationId], [ProductVersion])
    VALUES (N'20260528072311_AddSpsApprovalWorkflow', N'8.0.0');
END;

");
    }
    else if (provider.Contains("Sqlite", StringComparison.OrdinalIgnoreCase))
    {
        context.Database.ExecuteSqlRaw(@"
CREATE TABLE IF NOT EXISTS ApprovalRequests (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    RequestCode TEXT NOT NULL,
    ActionType INTEGER NOT NULL,
    TargetKey TEXT NOT NULL,
    RequestComment TEXT NULL,
    ReturnUrl TEXT NULL,
    PayloadJson TEXT NULL,
    RequesterUserName TEXT NOT NULL,
    RequesterRole TEXT NOT NULL,
    ApproverUserName TEXT NULL,
    ApproverRole TEXT NULL,
    ApproverComment TEXT NULL,
    ApprovalToken TEXT NULL,
    TokenExpiresAt TEXT NULL,
    Status INTEGER NOT NULL,
    CreatedAt TEXT NOT NULL,
    UpdatedAt TEXT NULL,
    ReviewedAt TEXT NULL,
    ConsumedAt TEXT NULL
);

CREATE TABLE IF NOT EXISTS ApprovalRequestLogs (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    ApprovalRequestId INTEGER NOT NULL,
    FromStatus INTEGER NULL,
    ToStatus INTEGER NOT NULL,
    Comment TEXT NULL,
    ActorUserName TEXT NOT NULL,
    ActorRole TEXT NOT NULL,
    CreatedAt TEXT NOT NULL,
    FOREIGN KEY(ApprovalRequestId) REFERENCES ApprovalRequests(Id) ON DELETE CASCADE
);

CREATE UNIQUE INDEX IF NOT EXISTS IX_ApprovalRequests_RequestCode ON ApprovalRequests(RequestCode);
CREATE INDEX IF NOT EXISTS IX_ApprovalRequests_Status_CreatedAt ON ApprovalRequests(Status, CreatedAt);
CREATE INDEX IF NOT EXISTS IX_ApprovalRequests_RequesterUserName_Status ON ApprovalRequests(RequesterUserName, Status);
CREATE INDEX IF NOT EXISTS IX_ApprovalRequests_ActionType_TargetKey_RequesterUserName_Status ON ApprovalRequests(ActionType, TargetKey, RequesterUserName, Status);
CREATE INDEX IF NOT EXISTS IX_ApprovalRequestLogs_ApprovalRequestId ON ApprovalRequestLogs(ApprovalRequestId);
");
    }
}

static void EnsureSpsDocumentActivationColumn(ApplicationDbContext context)
{
    var provider = context.Database.ProviderName ?? string.Empty;

    if (provider.Contains("SqlServer", StringComparison.OrdinalIgnoreCase))
    {
        context.Database.ExecuteSqlRaw(@"
IF COL_LENGTH(N'[dbo].[SpsNoDocs]', N'IsActive') IS NULL
BEGIN
    ALTER TABLE [dbo].[SpsNoDocs]
    ADD [IsActive] BIT NOT NULL CONSTRAINT [DF_SpsNoDocs_IsActive] DEFAULT(1) WITH VALUES;
END;
");

        return;
    }

    if (provider.Contains("Sqlite", StringComparison.OrdinalIgnoreCase))
    {
        try
        {
            context.Database.ExecuteSqlRaw("ALTER TABLE SpsNoDocs ADD COLUMN IsActive INTEGER NOT NULL DEFAULT 1;");
        }
        catch (Exception ex) when (ex.Message.Contains("duplicate column name", StringComparison.OrdinalIgnoreCase))
        {
            // Column already exists; no action needed.
        }
    }
}

static void EnsureSpsItemListLengthColumns(ApplicationDbContext context)
{
    var provider = context.Database.ProviderName ?? string.Empty;

    if (provider.Contains("SqlServer", StringComparison.OrdinalIgnoreCase))
    {
        context.Database.ExecuteSqlRaw(@"
IF COL_LENGTH(N'[dbo].[SpsItemLists]', N'LengthShrinkMin') IS NULL
    ALTER TABLE [dbo].[SpsItemLists] ADD [LengthShrinkMin] REAL NULL;

IF COL_LENGTH(N'[dbo].[SpsItemLists]', N'LengthShrinkStd') IS NULL
    ALTER TABLE [dbo].[SpsItemLists] ADD [LengthShrinkStd] REAL NULL;

IF COL_LENGTH(N'[dbo].[SpsItemLists]', N'LengthShrinkMax') IS NULL
    ALTER TABLE [dbo].[SpsItemLists] ADD [LengthShrinkMax] REAL NULL;

IF COL_LENGTH(N'[dbo].[SpsItemLists]', N'LengthCuttingMin') IS NULL
    ALTER TABLE [dbo].[SpsItemLists] ADD [LengthCuttingMin] REAL NULL;

IF COL_LENGTH(N'[dbo].[SpsItemLists]', N'LengthCuttingStd') IS NULL
    ALTER TABLE [dbo].[SpsItemLists] ADD [LengthCuttingStd] REAL NULL;

IF COL_LENGTH(N'[dbo].[SpsItemLists]', N'LengthCuttingMax') IS NULL
    ALTER TABLE [dbo].[SpsItemLists] ADD [LengthCuttingMax] REAL NULL;
");
    }
    else if (provider.Contains("Sqlite", StringComparison.OrdinalIgnoreCase))
    {
        var statements = new[]
        {
            "ALTER TABLE SpsItemLists ADD COLUMN LengthShrinkMin REAL NULL;",
            "ALTER TABLE SpsItemLists ADD COLUMN LengthShrinkStd REAL NULL;",
            "ALTER TABLE SpsItemLists ADD COLUMN LengthShrinkMax REAL NULL;",
            "ALTER TABLE SpsItemLists ADD COLUMN LengthCuttingMin REAL NULL;",
            "ALTER TABLE SpsItemLists ADD COLUMN LengthCuttingStd REAL NULL;",
            "ALTER TABLE SpsItemLists ADD COLUMN LengthCuttingMax REAL NULL;"
        };
        foreach (var sql in statements)
        {
            try
            {
                context.Database.ExecuteSqlRaw(sql);
            }
            catch (Exception ex) when (ex.Message.Contains("duplicate column name", StringComparison.OrdinalIgnoreCase))
            {
                // Already exists
            }
        }
    }
}

static void EnsureProductionReportExtendedColumns(ApplicationDbContext context)
{
    var provider = context.Database.ProviderName ?? string.Empty;

    if (provider.Contains("SqlServer", StringComparison.OrdinalIgnoreCase))
    {
        context.Database.ExecuteSqlRaw(@"
-- ProductionReports new columns
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitAmMeterInner') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitAmMeterInner] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitHeadTempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitHeadTempMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitCylinder1TempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitCylinder1TempMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitCylinder2TempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitCylinder2TempMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitCylinder3TempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitCylinder3TempMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitScrewTempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitScrewTempMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitScrewSpeedMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitScrewSpeedMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitFeedRollRatioMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitFeedRollRatioMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitPressureMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitPressureMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitAmMeterMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitAmMeterMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitAmMeterOuter') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitAmMeterOuter] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitOdSensor') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitOdSensor] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitCoolConveyorSpeed2') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitCoolConveyorSpeed2] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'InitCuttingSpeed') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [InitCuttingSpeed] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'AlarmFunctionCheck') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [AlarmFunctionCheck] NVARCHAR(50) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'LengthShrinkMin') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [LengthShrinkMin] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'LengthShrinkStd') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [LengthShrinkStd] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'LengthShrinkMax') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [LengthShrinkMax] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'LengthCuttingMin') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [LengthCuttingMin] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'LengthCuttingStd') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [LengthCuttingStd] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'LengthCuttingMax') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [LengthCuttingMax] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'ActualLengthShrink') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [ActualLengthShrink] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReports]', N'ActualLengthCutting') IS NULL
    ALTER TABLE [dbo].[ProductionReports] ADD [ActualLengthCutting] DECIMAL(18,2) NULL;

-- ProductionReadings new columns
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'AmMeterInner') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [AmMeterInner] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'HeadTempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [HeadTempMiddle] INT NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'Cylinder1TempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [Cylinder1TempMiddle] INT NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'Cylinder2TempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [Cylinder2TempMiddle] INT NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'Cylinder3TempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [Cylinder3TempMiddle] INT NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'ScrewTempMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [ScrewTempMiddle] INT NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'ScrewSpeedMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [ScrewSpeedMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'FeedRollRatioMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [FeedRollRatioMiddle] INT NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'PressureMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [PressureMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'AmMeterMiddle') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [AmMeterMiddle] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'AmMeterOuter') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [AmMeterOuter] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'OdSensor') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [OdSensor] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'CoolConveyorSpeed2') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [CoolConveyorSpeed2] DECIMAL(18,2) NULL;
IF COL_LENGTH(N'[dbo].[ProductionReadings]', N'CuttingSpeed') IS NULL
    ALTER TABLE [dbo].[ProductionReadings] ADD [CuttingSpeed] DECIMAL(18,2) NULL;
");
    }
    else if (provider.Contains("Sqlite", StringComparison.OrdinalIgnoreCase))
    {
        var statements = new[]
        {
            "ALTER TABLE ProductionReports ADD COLUMN InitAmMeterInner REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitHeadTempMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitCylinder1TempMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitCylinder2TempMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitCylinder3TempMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitScrewTempMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitScrewSpeedMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitFeedRollRatioMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitPressureMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitAmMeterMiddle REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitAmMeterOuter REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitOdSensor REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitCoolConveyorSpeed2 REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN InitCuttingSpeed REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN AlarmFunctionCheck TEXT NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN LengthShrinkMin REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN LengthShrinkStd REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN LengthShrinkMax REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN LengthCuttingMin REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN LengthCuttingStd REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN LengthCuttingMax REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN ActualLengthShrink REAL NULL;",
            "ALTER TABLE ProductionReports ADD COLUMN ActualLengthCutting REAL NULL;",

            "ALTER TABLE ProductionReadings ADD COLUMN AmMeterInner REAL NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN HeadTempMiddle INTEGER NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN Cylinder1TempMiddle INTEGER NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN Cylinder2TempMiddle INTEGER NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN Cylinder3TempMiddle INTEGER NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN ScrewTempMiddle INTEGER NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN ScrewSpeedMiddle REAL NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN FeedRollRatioMiddle INTEGER NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN PressureMiddle REAL NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN AmMeterMiddle REAL NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN AmMeterOuter REAL NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN OdSensor REAL NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN CoolConveyorSpeed2 REAL NULL;",
            "ALTER TABLE ProductionReadings ADD COLUMN CuttingSpeed REAL NULL;"
        };
        foreach (var sql in statements)
        {
            try
            {
                context.Database.ExecuteSqlRaw(sql);
            }
            catch (Exception ex) when (ex.Message.Contains("duplicate column name", StringComparison.OrdinalIgnoreCase))
            {
                // Already exists
            }
        }
    }
}

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
    app.UseHttpsRedirection();
}
app.UseStaticFiles();

// Tangkap error koneksi DB (saat server tidak bisa dijangkau / wifi mati)
app.UseDatabaseOfflineFallback();

app.UseRouting();
app.UseWebSockets();

app.UseSession();
app.UseAuthorization();

app.Map("/ws/dashboard", async context =>
{
    if (context.WebSockets.IsWebSocketRequest)
    {
        using var webSocket = await context.WebSockets.AcceptWebSocketAsync();
        await DashboardWebSocketHandler.HandleAsync(webSocket, context.RequestServices);
    }
    else
    {
        context.Response.StatusCode = StatusCodes.Status400BadRequest;
    }
});

app.MapHub<DashboardHub>("/hub/dashboard");

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.Run();
