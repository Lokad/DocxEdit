using Lokad.DocxEdit;

internal sealed class FileSystemCommandHost : IDocxCommandHost
{
    public TextWriter StandardOutput => Console.Out;
    public TextWriter StandardError => Console.Error;

    public bool PathsEqual(string first, string second) =>
        string.Equals(Path.GetFullPath(first), Path.GetFullPath(second),
            OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

    public string CombinePath(string directory, string fileName) => Path.Combine(directory, fileName);

    public ValueTask<Stream> OpenReadAsync(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return ValueTask.FromResult<Stream>(path == "-" ? Console.OpenStandardInput() : OpenFile(path));
    }

    public ValueTask<TextReader> OpenTextAsync(string path, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return ValueTask.FromResult<TextReader>(path == "-"
            ? new BorrowedTextReader(Console.In)
            : new StreamReader(OpenFile(path)));
    }

    private static FileStream OpenFile(string path) =>
        new(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.Asynchronous | FileOptions.SequentialScan);

    public async ValueTask PublishFileAsync(string path, Stream contents, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        string destination = Path.GetFullPath(path);
        string directory = Path.GetDirectoryName(destination)!;
        Directory.CreateDirectory(directory);
        string temporary = Path.Combine(directory, Path.GetFileName(destination) + ".tmp-" + Guid.NewGuid().ToString("N") + ".docxedit-tmp");
        try
        {
            await using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None, 81920, FileOptions.Asynchronous))
            {
                await contents.CopyToAsync(output, cancellationToken).ConfigureAwait(false);
            }
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporary, destination, overwrite: true);
        }
        finally
        {
            try { File.Delete(temporary); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    public async ValueTask WriteStandardOutputAsync(Stream contents, CancellationToken cancellationToken)
    {
        using Stream output = Console.OpenStandardOutput();
        await contents.CopyToAsync(output, cancellationToken).ConfigureAwait(false);
    }

    public IDocxAssetProvider GetAssetProvider(string patchPath) =>
        new FileSystemAssetProvider(patchPath == "-" ? Directory.GetCurrentDirectory() : Path.GetDirectoryName(Path.GetFullPath(patchPath)));

    private sealed class BorrowedTextReader(TextReader inner) : TextReader
    {
        public override ValueTask<int> ReadAsync(Memory<char> buffer, CancellationToken cancellationToken = default) =>
            inner.ReadAsync(buffer, cancellationToken);
    }

    private sealed class FileSystemAssetProvider : IDocxAssetProvider
    {
        private readonly string? _baseDirectory;

        public FileSystemAssetProvider(string? baseDirectory)
        {
            _baseDirectory = baseDirectory;
        }

        public bool TryOpen(
            string reference,
            out Stream stream,
            out string? contentTypeHint,
            out string? fileNameHint)
        {
            stream = Stream.Null;
            contentTypeHint = null;
            fileNameHint = null;

            try
            {
                if (_baseDirectory is not null && !Path.IsPathRooted(reference))
                {
                    string sibling = Path.GetFullPath(Path.Combine(_baseDirectory, reference));
                    if (File.Exists(sibling))
                    {
                        stream = File.OpenRead(sibling);
                        fileNameHint = Path.GetFileName(sibling);
                        return true;
                    }
                }

                string path = Path.GetFullPath(reference);
                if (!File.Exists(path))
                {
                    return false;
                }

                stream = File.OpenRead(path);
                fileNameHint = Path.GetFileName(path);
                return true;
            }
            catch (Exception ex) when (ex is ArgumentException or IOException or NotSupportedException or UnauthorizedAccessException)
            {
                return false;
            }
        }
    }


}
