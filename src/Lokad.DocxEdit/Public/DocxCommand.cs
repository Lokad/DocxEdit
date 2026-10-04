namespace Lokad.DocxEdit;

/// <summary>Runs the canonical docxedit commands against caller-provided I/O.</summary>
public static class DocxCommand
{
    /// <summary>Runs one invocation with default package quotas. Arguments exclude the executable name.</summary>
    public static Task<int> RunAsync(IReadOnlyList<string> arguments, IDocxCommandHost host, CancellationToken cancellationToken)
        => RunAsync(arguments, host, new DocxCommandOptions(), cancellationToken);

    /// <summary>Runs one invocation with host-selected quotas. Cancellation propagates; normal command failures return CLI exit codes.</summary>
    public static Task<int> RunAsync(IReadOnlyList<string> arguments, IDocxCommandHost host, DocxCommandOptions options, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        ArgumentNullException.ThrowIfNull(host);
        ArgumentNullException.ThrowIfNull(options);
        return new DocxCommandExecution(host, options, cancellationToken).ExecuteAsync(arguments.ToArray());
    }
}

/// <summary>Host-selected policy for an embedded command; command flags cannot raise these quotas.</summary>
public sealed class DocxCommandOptions
{
    /// <summary>Limits document loading, creation, and individual image assets.</summary>
    public DocxPackageLimits Quotas { get; init; } = DocxPackageLimits.Default;
}

/// <summary>I/O boundary for hosted commands. The library never accesses the machine filesystem or console.</summary>
/// <remarks>Each invocation borrows the host and its writers. Open methods return owned handles that the runner disposes. Path methods follow the host's namespace and working directory, including case sensitivity.</remarks>
public interface IDocxCommandHost
{
    /// <summary>Borrowed text stdout, also used for JSON. Never disposed by the command.</summary>
    TextWriter StandardOutput { get; }
    /// <summary>Borrowed text stderr. Never disposed by the command.</summary>
    TextWriter StandardError { get; }
    /// <summary>Compares two paths after resolving them within the host namespace.</summary>
    bool PathsEqual(string first, string second);
    /// <summary>Combines an export directory and a library-generated leaf filename within the host namespace.</summary>
    string CombinePath(string directory, string fileName);
    /// <summary>Opens binary input; '-' means standard input. The returned stream is owned by the command.</summary>
    ValueTask<Stream> OpenReadAsync(string path, CancellationToken cancellationToken);
    /// <summary>Opens text input; '-' means standard input. The returned reader is owned by the command.</summary>
    ValueTask<TextReader> OpenTextAsync(string path, CancellationToken cancellationToken);
    /// <summary>Publishes a complete staged file. Contents are borrowed at their current position until completion. Hosts must preserve any old destination if copying fails or is canceled and check cancellation before publication.</summary>
    ValueTask PublishFileAsync(string path, Stream contents, CancellationToken cancellationToken);
    /// <summary>Copies document bytes to binary stdout. Contents are borrowed until completion; stdout is never closed by the command. Output failures can leave a partial binary stream.</summary>
    ValueTask WriteStandardOutputAsync(Stream contents, CancellationToken cancellationToken);
    /// <summary>Returns a borrowed asynchronous provider bound to the patch's host location, or null if assets are unsupported. '-' denotes a stdin patch. This method creates no I/O; opens and reads happen through the provider with cancellation.</summary>
    IDocxAsyncAssetProvider? GetAssetProvider(string patchPath);
}
