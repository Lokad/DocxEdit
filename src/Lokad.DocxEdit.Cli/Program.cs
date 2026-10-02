using Lokad.DocxEdit;

return await ProgramMain.RunAsync(args, CancellationToken.None);

/// <summary>Local CLI adapter for the public hosted command runner.</summary>
public static class ProgramMain
{
    /// <summary>Compatibility entry point for synchronous command-line tests.</summary>
    public static int Run(string[] args) => RunAsync(args, CancellationToken.None).GetAwaiter().GetResult();

    /// <summary>Runs one command using the local filesystem and process standard streams.</summary>
    public static Task<int> RunAsync(string[] args, CancellationToken cancellationToken)
        => DocxCommand.RunAsync(args, new FileSystemCommandHost(), cancellationToken);
}
