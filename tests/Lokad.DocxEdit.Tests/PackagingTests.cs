using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace Lokad.DocxEdit.Tests;

public static class PackagingTests
{
    private const string PackageId = "Lokad.DocxEdit";
    private const string PackageVersion = "0.1.0";

    [Fact]
    public static void ReleaseNuGetPackWritesPackageAndSymbolsToArtifacts()
    {
        string repoRoot = FindRepoRoot();
        string artifactDirectory = Path.Combine(repoRoot, "artifacts", "nuget");
        Directory.CreateDirectory(artifactDirectory);
        string packagePath = Path.Combine(artifactDirectory, $"{PackageId}.{PackageVersion}.nupkg");
        string symbolPackagePath = Path.Combine(artifactDirectory, $"{PackageId}.{PackageVersion}.snupkg");
        DeleteIfExists(packagePath);
        DeleteIfExists(symbolPackagePath);

        PackResult result = RunDotnetPack(repoRoot, "Release");

        Assert.True(result.ExitCode == 0, result.Output + result.Error);
        Assert.True(File.Exists(packagePath), result.Output + result.Error);
        Assert.True(File.Exists(symbolPackagePath), result.Output + result.Error);

        using var archive = ZipFile.OpenRead(packagePath);
        Assert.NotNull(archive.GetEntry("lib/net10.0/Lokad.DocxEdit.dll"));
        Assert.NotNull(archive.GetEntry("lib/net10.0/Lokad.DocxEdit.xml"));
        Assert.NotNull(archive.GetEntry("README.md"));
        Assert.NotNull(archive.GetEntry("CHANGELOG.md"));
        Assert.NotNull(archive.GetEntry("LICENSE.txt"));
        Assert.NotNull(archive.GetEntry("icon.png"));

        ZipArchiveEntry nuspecEntry = Assert.Single(archive.Entries, entry => entry.FullName.EndsWith(".nuspec", StringComparison.Ordinal));
        using Stream stream = nuspecEntry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        string nuspec = reader.ReadToEnd();
        Assert.DoesNotContain("<dependency ", nuspec, StringComparison.OrdinalIgnoreCase);

        using var symbols = ZipFile.OpenRead(symbolPackagePath);
        Assert.NotNull(symbols.GetEntry("lib/net10.0/Lokad.DocxEdit.pdb"));
    }

    [Fact]
    public static void NonReleaseNuGetPackIsRejected()
    {
        string repoRoot = FindRepoRoot();
        using TempDirectory temp = TempDirectory.Create();

        PackResult result = RunDotnetPack(repoRoot, "Debug", "-o", temp.Path);

        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("NuGet packages must be produced with -c Release", result.Output + result.Error, StringComparison.Ordinal);
        Assert.Empty(Directory.EnumerateFiles(temp.Path, "*.nupkg"));
        Assert.Empty(Directory.EnumerateFiles(temp.Path, "*.snupkg"));
    }

    private static PackResult RunDotnetPack(string repoRoot, string configuration, params string[] extraArguments)
    {
        var startInfo = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = repoRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("pack");
        startInfo.ArgumentList.Add(Path.Combine(repoRoot, "src", "Lokad.DocxEdit", "Lokad.DocxEdit.csproj"));
        startInfo.ArgumentList.Add("-c");
        startInfo.ArgumentList.Add(configuration);
        startInfo.ArgumentList.Add("--nologo");
        foreach (string argument in extraArguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start dotnet pack.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();

        return new PackResult(process.ExitCode, output, error);
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private static string FindRepoRoot()
    {
        string? directory = AppContext.BaseDirectory;
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory, "Lokad.DocxEdit.slnx")))
            {
                return directory;
            }

            directory = Directory.GetParent(directory)?.FullName;
        }

        throw new InvalidOperationException("Could not locate repository root.");
    }

    private sealed record PackResult(int ExitCode, string Output, string Error);

    private sealed class TempDirectory : IDisposable
    {
        private TempDirectory(string path)
        {
            Path = path;
        }

        public string Path { get; }

        public static TempDirectory Create()
        {
            string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "docxedit-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return new TempDirectory(path);
        }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
