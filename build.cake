#tool "nuget:?package=NuGet.CommandLine&version=5.11.0"

#addin "nuget:?package=Cake.MinVer&version=1.0.1"
#addin "nuget:?package=Cake.Args&version=1.0.1"

var target       = ArgumentOrDefault<string>("target") ?? "pack";
var buildVersion = MinVer(s => s.WithTagPrefix("v").WithDefaultPreReleasePhase("preview"));

Task("clean")
    .Does(() =>
{
    CleanDirectories("./artifact/**");
    CleanDirectories("./packages/**");
    CleanDirectories("./**/^{bin,obj}");
});

Task("restore")
    .IsDependentOn("clean")
    .Does(() =>
{
    NuGetRestore("./serilog-sinks-exceldnalogdisplay.sln", new NuGetRestoreSettings
    {
        NoCache = true,
    });
});

Task("build")
    .IsDependentOn("restore")
    .DoesForEach(new[] { "Debug", "Release" }, (configuration) =>
{
    MSBuild("./serilog-sinks-exceldnalogdisplay.sln", settings => settings
        .SetConfiguration(configuration)
        .UseToolVersion(MSBuildToolVersion.VS2019)
        .WithTarget("Rebuild")
        .WithProperty("Version", buildVersion.Version)
        .WithProperty("FileVersion", buildVersion.FileVersion)
        .WithProperty("ContinuousIntegrationBuild", "true")
    );
});

Task("test")
    .IsDependentOn("build")
    .Does(() =>
{
    var settings = new DotNetCoreTestSettings
    {
        Configuration = "Release",
        NoRestore = true,
        NoBuild = true,
    };

    var projectFiles = GetFiles("./test/**/*.csproj");
    foreach (var file in projectFiles)
    {
        DotNetCoreTest(file.FullPath, settings);
    }
});

Task("pack")
    .IsDependentOn("test")
    .Does(() =>
{
    var releaseNotes = $"https://github.com/serilog-contrib/serilog-sinks-exceldnalogdisplay/releases/tag/v{buildVersion.Version}";

    DotNetCorePack("./src/Serilog.Sinks.ExcelDnaLogDisplay/Serilog.Sinks.ExcelDnaLogDisplay.csproj", new DotNetCorePackSettings
    {
        Configuration = "Release",
        NoRestore = true,
        NoBuild = true,
        IncludeSymbols = true,
        IncludeSource = true,
        OutputDirectory = "./artifact/nuget",
        MSBuildSettings = new DotNetCoreMSBuildSettings()
            .WithProperty("Version", buildVersion.Version)
            .WithProperty("PackageReleaseNotes", releaseNotes)
    });
});

Task("push")
    .IsDependentOn("pack")
    .Does(() =>
{
    var url =  EnvironmentVariable("NUGET_URL");
    if (string.IsNullOrWhiteSpace(url))
    {
        Information("No NuGet URL specified. Skipping publishing of NuGet packages");
        return;
    }

    var apiKey =  EnvironmentVariable("NUGET_API_KEY");
    if (string.IsNullOrWhiteSpace(apiKey))
    {
        Information("No NuGet API key specified. Skipping publishing of NuGet packages");
        return;
    }

    var nugetPushSettings = new DotNetCoreNuGetPushSettings
    {
        Source = url,
        ApiKey = apiKey,
    };

    foreach (var nugetPackageFile in GetFiles("./artifact/nuget/*.nupkg"))
    {
        DotNetCoreNuGetPush(nugetPackageFile.FullPath, nugetPushSettings);
    }
});

RunTarget(target);
