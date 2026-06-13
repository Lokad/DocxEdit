using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace Lokad.DocxEdit.Tests;

public static class PackagingTests
{
    [Fact]
    public static void NuGetPackageContainsExpectedFilesAndNoDependencies()
    {
        string repoRoot = FindRepoRoot();
        using TempDirectory temp = TempDirectory.Create();
        var startInfo = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = repoRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.ArgumentList.Add("pack");
        startInfo.ArgumentList.Add(Path.Combine(repoRoot, "src", "Lokad.DocxEdit", "Lokad.DocxEdit.csproj"));
        startInfo.ArgumentList.Add("-c");
        startInfo.ArgumentList.Add("Release");
        startInfo.ArgumentList.Add("-o");
        startInfo.ArgumentList.Add(temp.Path);
        startInfo.ArgumentList.Add("--nologo");

        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start dotnet pack.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();

        Assert.True(process.ExitCode == 0, output + error);
        string packagePath = Assert.Single(Directory.EnumerateFiles(temp.Path, "Lokad.DocxEdit.*.nupkg"));
        using var archive = ZipFile.OpenRead(packagePath);
        Assert.NotNull(archive.GetEntry("lib/net10.0/Lokad.DocxEdit.dll"));
        Assert.NotNull(archive.GetEntry("README.md"));
        Assert.NotNull(archive.GetEntry("CHANGELOG.md"));
        Assert.NotNull(archive.GetEntry("LICENSE.txt"));

        ZipArchiveEntry nuspecEntry = Assert.Single(archive.Entries, entry => entry.FullName.EndsWith(".nuspec", StringComparison.Ordinal));
        using Stream stream = nuspecEntry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        string nuspec = reader.ReadToEnd();
        Assert.DoesNotContain("<dependency ", nuspec, StringComparison.OrdinalIgnoreCase);
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
