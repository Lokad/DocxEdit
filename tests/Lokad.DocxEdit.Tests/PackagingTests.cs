using System.Diagnostics;
using System.IO.Compression;
using System.Reflection.Metadata;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;

using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

public static class PackagingTests
{
    private const string PackageId = "Lokad.DocxEdit";

    private static string ReadProjectVersion(string repoRoot)
    {
        string csprojPath = Path.Combine(repoRoot, "src", "Lokad.DocxEdit", "Lokad.DocxEdit.csproj");
        var document = XDocument.Load(csprojPath);
        string? version = document.Descendants().FirstOrDefault(static element => element.Name.LocalName == "Version")?.Value?.Trim();
        Assert.False(string.IsNullOrWhiteSpace(version), "Expected a Version element in Lokad.DocxEdit.csproj.");
        string? projectId = document.Descendants().FirstOrDefault(static element => element.Name.LocalName == "PackageId")?.Value?.Trim();
        Assert.Equal(PackageId, projectId);
        return version!;
    }

    [Fact]
    public static void ReleaseNuGetPackWritesPackageAndSymbolsToArtifacts()
    {
        string repoRoot = FindRepoRoot();
        string expectedVersion = ReadProjectVersion(repoRoot);
        string artifactDirectory = Path.Combine(repoRoot, "artifacts", "nuget");
        Directory.CreateDirectory(artifactDirectory);
        string packagePath = Path.Combine(artifactDirectory, $"{PackageId}.{expectedVersion}.nupkg");
        string symbolPackagePath = Path.Combine(artifactDirectory, $"{PackageId}.{expectedVersion}.snupkg");
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
        string nuspec;
        using (Stream stream = nuspecEntry.Open())
        using (var reader = new StreamReader(stream, Encoding.UTF8))
        {
            nuspec = reader.ReadToEnd();
        }

        string expectedCommit = GetGitHead(repoRoot);
        VerifyNuspec(nuspec, expectedVersion, expectedCommit);
        VerifyPortablePdb(symbolPackagePath, expectedCommit);
        VerifyPackageConsumerFromLocalFeed(repoRoot, packagePath);
    }

    private static void VerifyNuspec(string nuspec, string expectedVersion, string expectedCommit)
    {
        var document = XDocument.Parse(nuspec);
        string? id = document.Descendants().FirstOrDefault(static element => element.Name.LocalName == "id")?.Value?.Trim();
        Assert.Equal(PackageId, id);
        string? version = document.Descendants().FirstOrDefault(static element => element.Name.LocalName == "version")?.Value?.Trim();
        Assert.Equal(expectedVersion, version);
        var repository = document.Descendants().FirstOrDefault(static element => element.Name.LocalName == "repository");
        Assert.NotNull(repository);
        Assert.Equal("git", repository.Attribute("type")?.Value);
        string? url = repository.Attribute("url")?.Value;
        Assert.NotNull(url);
        Assert.Contains("github.com/Lokad/DocxEdit", url, StringComparison.Ordinal);
        string? commit = repository.Attribute("commit")?.Value;
        Assert.False(string.IsNullOrWhiteSpace(commit), "Expected a repository commit in the nuspec.");
        Assert.Equal(expectedCommit, commit);
        Assert.Equal(40, commit!.Length);
        try
        {
            byte[] commitBytes = Convert.FromHexString(commit);
            Assert.Equal(20, commitBytes.Length);
        }
        catch (FormatException exception)
        {
            Assert.Fail("Expected a 40-character hexadecimal commit in the nuspec: " + exception.Message);
        }

        Assert.False(document.Descendants().Any(static element => element.Name.LocalName == "dependency"), "Expected no runtime package dependencies in the nuspec.");
        var groups = document.Descendants().Where(static element => element.Name.LocalName == "group").ToList();
        Assert.NotEmpty(groups);
        Assert.All(groups, static group => Assert.Contains("net10.0", group.Attribute("targetFramework")?.Value ?? string.Empty, StringComparison.Ordinal));
    }

    private static string GetGitHead(string repoRoot)
    {
        PackResult git = RunProcess("git", ["rev-parse", "HEAD"], repoRoot, null, TimeSpan.FromSeconds(30));
        Assert.True(git.ExitCode == 0, git.Output + git.Error);
        string head = git.Output.Trim();
        Assert.Equal(40, head.Length);
        try
        {
            Assert.Equal(20, Convert.FromHexString(head).Length);
        }
        catch (FormatException exception)
        {
            Assert.Fail("Expected git rev-parse HEAD to return hexadecimal: " + exception.Message);
        }

        return head;
    }

    private static void VerifyPortablePdb(string symbolPackagePath, string expectedCommit)
    {
        byte[] pdbBytes;
        using (var archive = ZipFile.OpenRead(symbolPackagePath))
        {
            ZipArchiveEntry? pdbEntry = archive.GetEntry("lib/net10.0/Lokad.DocxEdit.pdb");
            Assert.NotNull(pdbEntry);
            using Stream pdbStream = pdbEntry.Open();
            using var pdbMemory = new MemoryStream();
            pdbStream.CopyTo(pdbMemory);
            pdbBytes = pdbMemory.ToArray();
        }

        Assert.True(pdbBytes.Length > 0, "Expected PDB bytes in the symbol package.");
        using var pdbStream2 = new MemoryStream(pdbBytes);
        using var provider = MetadataReaderProvider.FromPortablePdbStream(pdbStream2);
        MetadataReader reader = provider.GetMetadataReader();
        var documentNames = new List<string>();
        foreach (DocumentHandle handle in reader.Documents)
        {
            Document document = reader.GetDocument(handle);
            string name = reader.GetString(document.Name);
            documentNames.Add(name);
        }

        Assert.NotEmpty(documentNames);
        Assert.All(documentNames, static name => Assert.StartsWith("/_/", name, StringComparison.Ordinal));
        Assert.All(documentNames, static name => Assert.DoesNotContain("\\", name, StringComparison.Ordinal));
        Assert.All(documentNames, static name => Assert.DoesNotContain(":", name, StringComparison.Ordinal));
        string? sourceLink = FindSourceLinkJson(reader);
        Assert.False(string.IsNullOrWhiteSpace(sourceLink), "Expected SourceLink JSON in the PDB.");
        using JsonDocument linkDocument = JsonDocument.Parse(sourceLink!);
        Assert.True(linkDocument.RootElement.TryGetProperty("documents", out JsonElement documents), "Expected a documents object in SourceLink JSON.");
        bool foundMapping = false;
        foreach (JsonProperty property in documents.EnumerateObject())
        {
            string target = property.Value.GetString() ?? string.Empty;
            Assert.Contains(expectedCommit, target, StringComparison.Ordinal);
            if (property.Name == "/_/*" && target.Contains("raw.githubusercontent.com/Lokad/DocxEdit/" + expectedCommit + "/", StringComparison.Ordinal))
            {
                foundMapping = true;
            }
        }

        Assert.True(foundMapping, "Expected a /_/* SourceLink mapping for commit " + expectedCommit + ".");
    }

    private static string? FindSourceLinkJson(MetadataReader reader)
    {
        foreach (CustomDebugInformationHandle handle in reader.CustomDebugInformation)
        {
            CustomDebugInformation info = reader.GetCustomDebugInformation(handle);
            byte[] bytes = reader.GetBlobBytes(info.Value);
            string text = Encoding.UTF8.GetString(bytes);
            if (text.Contains("documents", StringComparison.Ordinal) && text.Contains("githubusercontent", StringComparison.Ordinal))
            {
                return text;
            }
        }

        return null;
    }

    private static void VerifyPackageConsumerFromLocalFeed(string repoRoot, string packagePath)
    {
        string expectedVersion = ReadProjectVersion(repoRoot);
        string artifactDirectory = Path.Combine(repoRoot, "artifacts", "nuget");
        string feedPath = artifactDirectory.Replace("\\", "/");
        using TempDirectory consumerRoot = TempDirectory.Create();
        using TempDirectory packagesRoot = TempDirectory.Create();
        string projectPath = Path.Combine(consumerRoot.Path, "Consumer.csproj");
        string programPath = Path.Combine(consumerRoot.Path, "Program.cs");
        string configPath = Path.Combine(consumerRoot.Path, "NuGet.config");
        string consumerCsproj = $"<Project Sdk=\"Microsoft.NET.Sdk\">\n  <PropertyGroup>\n    <OutputType>Exe</OutputType>\n    <TargetFramework>net10.0</TargetFramework>\n    <Nullable>enable</Nullable>\n    <TreatWarningsAsErrors>true</TreatWarningsAsErrors>\n  </PropertyGroup>\n  <ItemGroup>\n    <PackageReference Include=\"{PackageId}\" Version=\"{expectedVersion}\" />\n  </ItemGroup>\n</Project>\n";
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
        string repoSdk = GetDotnetVersion(repoRoot);
        WriteConsumerSdkPin(consumerRoot.Path, repoSdk);
        string consumerSdk = GetDotnetVersion(consumerRoot.Path);
        Assert.Equal(repoSdk, consumerSdk);
        PackResult consumerResult = RunDotnetConsumer(consumerRoot.Path, packagesRoot.Path);
        Assert.True(consumerResult.ExitCode == 0, consumerResult.Output + consumerResult.Error);
        Assert.Contains("CONSUMER-OK", consumerResult.Output, StringComparison.Ordinal);
    }

    private const string ConsumerProgram = """
        using Lokad.DocxEdit;
        using System;
        using System.IO;
        using System.Threading;
        using System.Threading.Tasks;
        using System.Text.Json;
        sealed class FileAsset : IDocxAssetProvider, IDocxAsyncAssetProvider
        {
            public ValueTask<DocxAsset?> OpenAsync(string reference, CancellationToken cancellationToken)
            {
                cancellationToken.ThrowIfCancellationRequested();
                return ValueTask.FromResult(TryOpen(reference, out Stream stream, out string? type, out string? file)
                    ? new DocxAsset(stream, type, file) : null);
            }
            public bool TryOpen(string reference, out Stream stream, out string? contentTypeHint, out string? fileNameHint)
            {
                contentTypeHint = null;
                fileNameHint = null;
                try
                {
                    stream = File.OpenRead(reference);
                    return true;
                }
                catch
                {
                    stream = Stream.Null;
                    return false;
                }
            }
        }
        sealed class CommandHost : IDocxCommandHost
        {
            public StringWriter Output { get; } = new();
            public TextWriter StandardOutput => Output;
            public TextWriter StandardError { get; } = new StringWriter();
            public bool PathsEqual(string first, string second) => Path.GetFullPath(first) == Path.GetFullPath(second);
            public string CombinePath(string directory, string name) => Path.Combine(directory, name);
            public IDocxAsyncAssetProvider GetAssetProvider(string patchPath) => new FileAsset();
            public ValueTask<Stream> OpenReadAsync(string path, CancellationToken token) => ValueTask.FromResult<Stream>(File.OpenRead(path));
            public ValueTask<TextReader> OpenTextAsync(string path, CancellationToken token) => ValueTask.FromResult<TextReader>(File.OpenText(path));
            public ValueTask WriteStandardOutputAsync(Stream contents, CancellationToken token) => throw new NotSupportedException();
            public async ValueTask PublishFileAsync(string path, Stream contents, CancellationToken token)
            {
                string temporary = path + ".staged";
                try
                {
                    await using (var output = File.Create(temporary))
                        await contents.CopyToAsync(output, token);
                    token.ThrowIfCancellationRequested();
                    File.Move(temporary, path, overwrite: true);
                }
                finally { File.Delete(temporary); }
            }
        }
        static class Consumer
        {
            static int Main()
            {
                try
                {
                    using var textInput = File.OpenRead("seed-text.docx");
                    using var textPatch = File.OpenText("text.patch");
                    using var textStaged = new MemoryStream();
                    DocxApplyResult textResult = new DocxEditor().Apply(textInput, textPatch, textStaged);
                    if (!textResult.Success)
                    {
                        Console.Error.WriteLine("text apply failed");
                        return 1;
                    }
                    textStaged.Position = 0;
                    DocxReadResult textRead = new DocxEditor().Read(textStaged);
                    if (!textRead.Success || textRead.Paragraphs.Count != 1 || textRead.Paragraphs[0].Text != "Omega")
                    {
                        Console.Error.WriteLine("text readback mismatch");
                        return 1;
                    }
                    using var imageInput = File.OpenRead("seed-image.docx");
                    DocxMediaExtractResult before = new DocxEditor().ExtractMedia(imageInput);
                    if (!before.Success || before.Files.Count != 1)
                    {
                        Console.Error.WriteLine("extract failed");
                        return 1;
                    }
                    using var patchInput = File.OpenRead("seed-image.docx");
                    using var imagePatch = File.OpenText("image.patch");
                    using var imageStaged = new MemoryStream();
                    var options = new DocxEditOptions { AssetProvider = new FileAsset() };
                    DocxApplyResult imageResult = new DocxEditor().Apply(patchInput, imagePatch, imageStaged, options);
                    if (!imageResult.Success)
                    {
                        Console.Error.WriteLine("image apply failed");
                        return 1;
                    }
                    imageStaged.Position = 0;
                    DocxMediaExtractResult after = new DocxEditor().ExtractMedia(imageStaged);
                    if (!after.Success || after.Files.Count != 1)
                    {
                        Console.Error.WriteLine("re-extract failed");
                        return 1;
                    }
                    byte[] expected = File.ReadAllBytes("new.png");
                    if (!after.Files[0].Content.AsSpan().SequenceEqual(expected))
                    {
                        Console.Error.WriteLine("image bytes mismatch");
                        return 1;
                    }
                    var host = new CommandHost();
                    int commandExit = DocxCommand.RunAsync(
                        new[] { "apply", "seed-image.docx", "image.patch", "-o", "hosted.docx", "--json" },
                        host, CancellationToken.None).GetAwaiter().GetResult();
                    if (commandExit != 0) throw new Exception("hosted command failed: " + host.StandardError.ToString());
                    using var commandJson = JsonDocument.Parse(host.Output.ToString());
                    if (commandJson.RootElement.GetProperty("TrackChanges").GetString() != "off")
                        throw new Exception("hosted JSON wire format mismatch");
                    using var hostedInput = File.OpenRead("hosted.docx");
                    var hostedImages = new DocxEditor().ExtractMedia(hostedInput);
                    if (!hostedImages.Success || hostedImages.Files.Count != 1 || !hostedImages.Files[0].Content.AsSpan().SequenceEqual(expected))
                        throw new Exception("hosted asset round trip mismatch");
                    var jsonOptions = DocxJson.CreateOptions(false);
                    if (JsonSerializer.Serialize(TrackChangesMode.Preserve, jsonOptions) != "\"preserve\"")
                        throw new Exception("public JSON options mismatch");
                    Console.WriteLine("CONSUMER-OK");
                    return 0;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex.Message);
                    return 1;
                }
            }
        }
        """;

    private static void WriteConsumerSdkPin(string consumerDirectory, string sdkVersion)
    {
        string pin = "{\"sdk\":{\"version\":\"" + sdkVersion + "\",\"rollForward\":\"disable\"}}";
        File.WriteAllText(Path.Combine(consumerDirectory, "global.json"), pin);
    }

    private static string GetDotnetVersion(string workingDirectory)
    {
        PackResult version = RunProcess("dotnet", ["--version"], workingDirectory, null, TimeSpan.FromSeconds(60));
        Assert.True(version.ExitCode == 0, version.Output + version.Error);
        string sdk = version.Output.Trim();
        Assert.False(string.IsNullOrWhiteSpace(sdk), "Expected dotnet --version to return a version.");
        return sdk;
    }

    private static PackResult RunDotnetConsumer(string workingDirectory, string packagesDirectory)
    {
        var environment = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["NUGET_PACKAGES"] = packagesDirectory
        };
        return RunProcess("dotnet", ["run", "--nologo"], workingDirectory, environment, TimeSpan.FromSeconds(120));
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
        var arguments = new List<string>
        {
            "pack",
            Path.Combine(repoRoot, "src", "Lokad.DocxEdit", "Lokad.DocxEdit.csproj"),
            "-c",
            configuration,
            "--nologo"
        };
        arguments.AddRange(extraArguments);
        return RunProcess("dotnet", arguments, repoRoot, null, TimeSpan.FromMinutes(5));
    }

    private static PackResult RunProcess(string fileName, IReadOnlyList<string> arguments, string workingDirectory, IDictionary<string, string>? extraEnvironment, TimeSpan timeout)
    {
        var startInfo = new ProcessStartInfo(fileName)
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        foreach (string argument in arguments)
        {
            startInfo.ArgumentList.Add(argument);
        }

        if (extraEnvironment is not null)
        {
            foreach (var entry in extraEnvironment)
            {
                startInfo.Environment[entry.Key] = entry.Value;
            }
        }

        using Process process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start " + fileName + ".");
        Task<string> outputTask = process.StandardOutput.ReadToEndAsync();
        Task<string> errorTask = process.StandardError.ReadToEndAsync();
        bool exited = process.WaitForExit((int)timeout.TotalMilliseconds);
        if (!exited)
        {
            try
            {
                process.Kill(true);
            }
            catch (InvalidOperationException)
            {
            }
            catch (NotSupportedException)
            {
            }

            bool stopped = process.WaitForExit(10000);
            string timedOutOutput = DrainWithBound(outputTask, 10);
            string timedOutError = DrainWithBound(errorTask, 10);
            string detail = "Timed out after " + timeout.TotalSeconds + "s waiting for " + fileName + ". Exited after kill: " + stopped + ". Output: " + timedOutOutput + " Error: " + timedOutError;
            return new PackResult(-1, timedOutOutput, detail);
        }

        string finalOutput = DrainWithBound(outputTask, 10);
        string finalError = DrainWithBound(errorTask, 10);
        return new PackResult(process.ExitCode, finalOutput, finalError);
    }

    private static string DrainWithBound(Task<string> task, int seconds)
    {
        bool done = task.Wait(TimeSpan.FromSeconds(seconds));
        if (done)
        {
            return task.GetAwaiter().GetResult();
        }

        return "<output drain timed out after " + seconds + "s>";
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
