using System.Diagnostics;
using System.IO.Compression;
using System.Text;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

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

        Assert.Contains($"<id>{PackageId}</id>", nuspec, StringComparison.Ordinal);
        Assert.Contains($"<version>{PackageVersion}</version>", nuspec, StringComparison.Ordinal);
        Assert.Contains("commit=", nuspec, StringComparison.Ordinal);
        using var symbols = ZipFile.OpenRead(symbolPackagePath);
        ZipArchiveEntry? pdbEntry = symbols.GetEntry("lib/net10.0/Lokad.DocxEdit.pdb");
        Assert.NotNull(pdbEntry);
        byte[] pdbBytes;
        using (Stream pdbStream = pdbEntry.Open())
        using (var pdbMemory = new MemoryStream())
        {
            pdbStream.CopyTo(pdbMemory);
            pdbBytes = pdbMemory.ToArray();
        }
        string pdbText = Encoding.ASCII.GetString(pdbBytes);
        Assert.Contains("/_/", pdbText, StringComparison.Ordinal);
        VerifyPackageConsumerFromLocalFeed(repoRoot, packagePath);
    }

    private static void VerifyPackageConsumerFromLocalFeed(string repoRoot, string packagePath)
    {
        string artifactDirectory = Path.Combine(repoRoot, "artifacts", "nuget");
        string feedPath = artifactDirectory.Replace("\\", "/");
        using TempDirectory consumerRoot = TempDirectory.Create();
        using TempDirectory packagesRoot = TempDirectory.Create();
        string projectPath = Path.Combine(consumerRoot.Path, "Consumer.csproj");
        string programPath = Path.Combine(consumerRoot.Path, "Program.cs");
        string configPath = Path.Combine(consumerRoot.Path, "NuGet.config");
        string consumerCsproj = $"<Project Sdk=\"Microsoft.NET.Sdk\">\n  <PropertyGroup>\n    <OutputType>Exe</OutputType>\n    <TargetFramework>net10.0</TargetFramework>\n    <Nullable>enable</Nullable>\n    <TreatWarningsAsErrors>true</TreatWarningsAsErrors>\n  </PropertyGroup>\n  <ItemGroup>\n    <PackageReference Include=\"{PackageId}\" Version=\"{PackageVersion}\" />\n  </ItemGroup>\n</Project>\n";
        File.WriteAllText(projectPath, consumerCsproj);
        File.WriteAllText(programPath, ConsumerProgram);
        string nugetConfig = $"<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<configuration>\n  <packageSources>\n    <clear />\n    <add key=\"local\" value=\"{feedPath}\" />\n  </packageSources>\n</configuration>\n";
        File.WriteAllText(configPath, nugetConfig);
        using MemoryStream textSeed = CreateDocx("Alpha");
        File.WriteAllBytes(Path.Combine(consumerRoot.Path, "seed-text.docx"), textSeed.ToArray());
        using MemoryStream imageSeed = CreateDocxWithImage("png", "image/png", "old-png");
        File.WriteAllBytes(Path.Combine(consumerRoot.Path, "seed-image.docx"), imageSeed.ToArray());
        byte[] replacement = CreatePngBytes(4, 3);
        File.WriteAllBytes(Path.Combine(consumerRoot.Path, "new.png"), replacement);
        File.WriteAllText(Path.Combine(consumerRoot.Path, "text.patch"), "docxpatch 1\n\nop replace-text\ntarget M.P0001\nexpect-text Alpha\nfind Alpha\nwith Omega\nend\n");
        File.WriteAllText(Path.Combine(consumerRoot.Path, "image.patch"), "docxpatch 1\n\nop replace-image\ntarget M.I0001\nexpect-content-type image/png\nasset new.png\nend\n");
        PackResult consumerResult = RunDotnetConsumer(consumerRoot.Path, packagesRoot.Path);
        Assert.True(consumerResult.ExitCode == 0, consumerResult.Output + consumerResult.Error);
        Assert.Contains("CONSUMER-OK", consumerResult.Output, StringComparison.Ordinal);
    }

    private const string ConsumerProgram = "using Lokad.DocxEdit;\nusing System;\nusing System.IO;\nsealed class FileAsset : IDocxAssetProvider\n{\n    public bool TryOpen(string reference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)\n    {\n        contentTypeHint = null;\n        fileNameHint = null;\n        try\n        {\n            stream = File.OpenRead(reference);\n            return true;\n        }\n        catch\n        {\n            stream = Stream.Null;\n            return false;\n        }\n    }\n}\nstatic class Consumer\n{\n    static int Main()\n    {\n        try\n        {\n            using var textInput = File.OpenRead(\"seed-text.docx\");\n            using var textPatch = File.OpenText(\"text.patch\");\n            using var textStaged = new MemoryStream();\n            DocxApplyResult textResult = new DocxEditor().Apply(textInput, textPatch, textStaged);\n            if (!textResult.Success)\n            {\n                Console.Error.WriteLine(\"text apply failed\");\n                return 1;\n            }\n            textStaged.Position = 0;\n            DocxReadResult textRead = new DocxEditor().Read(textStaged);\n            if (!textRead.Success || textRead.Paragraphs.Count != 1 || textRead.Paragraphs[0].Text != \"Omega\")\n            {\n                Console.Error.WriteLine(\"text readback mismatch\");\n                return 1;\n            }\n            using var imageInput = File.OpenRead(\"seed-image.docx\");\n            DocxMediaExtractResult before = new DocxEditor().ExtractMedia(imageInput);\n            if (!before.Success || before.Files.Count != 1)\n            {\n                Console.Error.WriteLine(\"extract failed\");\n                return 1;\n            }\n            using var patchInput = File.OpenRead(\"seed-image.docx\");\n            using var imagePatch = File.OpenText(\"image.patch\");\n            using var imageStaged = new MemoryStream();\n            var options = new DocxEditOptions { AssetProvider = new FileAsset() };\n            DocxApplyResult imageResult = new DocxEditor().Apply(patchInput, imagePatch, imageStaged, options);\n            if (!imageResult.Success)\n            {\n                Console.Error.WriteLine(\"image apply failed\");\n                return 1;\n            }\n            imageStaged.Position = 0;\n            DocxMediaExtractResult after = new DocxEditor().ExtractMedia(imageStaged);\n            if (!after.Success || after.Files.Count != 1)\n            {\n                Console.Error.WriteLine(\"re-extract failed\");\n                return 1;\n            }\n            byte[] expected = File.ReadAllBytes(\"new.png\");\n            if (!after.Files[0].Content.AsSpan().SequenceEqual(expected))\n            {\n                Console.Error.WriteLine(\"image bytes mismatch\");\n                return 1;\n            }\n            Console.WriteLine(\"CONSUMER-OK\");\n            return 0;\n        }\n        catch (Exception ex)\n        {\n            Console.Error.WriteLine(ex.Message);\n            return 1;\n        }\n    }\n}\n";

    private static PackResult RunDotnetConsumer(string workingDirectory, string packagesDirectory)
    {
        var startInfo = new ProcessStartInfo("dotnet")
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        startInfo.Environment["NUGET_PACKAGES"] = packagesDirectory;
        startInfo.ArgumentList.Add("run");
        startInfo.ArgumentList.Add("--nologo");
        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start dotnet run.");
        Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
        Task<string> errorTask = process.StandardError.ReadToEndAsync();
        process.WaitForExit(120000);
        string output = outputTask.GetAwaiter().GetResult();
        string error = errorTask.GetAwaiter().GetResult();
        return new PackResult(process.ExitCode, output, error);
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
        // Drain both streams concurrently: sequential ReadToEnd calls deadlock
        // once the child fills the pipe nobody is reading.
        Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
        Task<string> errorTask = process.StandardError.ReadToEndAsync();
        process.WaitForExit();
        string output = outputTask.GetAwaiter().GetResult();
        string error = errorTask.GetAwaiter().GetResult();

        return new PackResult(process.ExitCode, output, error);
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }


    private sealed record PackResult(int ExitCode, string Output, string Error);

}
