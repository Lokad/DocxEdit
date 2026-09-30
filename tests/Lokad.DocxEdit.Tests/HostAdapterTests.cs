using static Lokad.DocxEdit.Tests.DocxTestFixtures;

namespace Lokad.DocxEdit.Tests;

// Portable host-adapter reference for the PLAN host-integration item: a version-gated,
// staged-output embedding of the public DocxEditor surface. Documents live behind
// opaque handles as immutable versioned byte snapshots; every check/apply names the
// expected input version, outputs stage to a buffer, and publication swaps bytes and
// bumps the version atomically. Failures, stale preconditions, and cancellation
// discard staged output, so the store never partially publishes. Cancellation is
// never masked: OperationCanceledException propagates with nothing published.
// All fixtures are synthetic.
public static class HostAdapterTests
{
    private sealed class StaleVersionException(long expected, long actual)
        : InvalidOperationException($"Stale input version {expected}; current version is {actual}.")
    {
        public long ExpectedVersion { get; } = expected;
        public long ActualVersion { get; } = actual;
    }

    private sealed class VersionedDocumentStore
    {
        private readonly Dictionary<string, List<byte[]>> versionsByHandle = new(StringComparer.Ordinal);

        public string Ingest(byte[] docx)
        {
            string handle = Guid.NewGuid().ToString("N");
            versionsByHandle[handle] = [(byte[])docx.Clone()];
            return handle;
        }

        public long CurrentVersion(string handle)
        {
            return Versions(handle).Count - 1;
        }

        public byte[] Snapshot(string handle, long version)
        {
            List<byte[]> versions = Versions(handle);
            return (byte[])versions[(int)version].Clone();
        }

        public DocxCheckResult Check(
            string handle,
            long expectedVersion,
            string patchText,
            DocxEditOptions? options = null,
            CancellationToken cancellationToken = default)
        {
            long current = GuardVersion(handle, expectedVersion);
            using var input = new MemoryStream(Snapshot(handle, current), writable: false);
            using var patch = new StringReader(patchText);
            return new DocxEditor().Check(input, patch, options ?? new DocxEditOptions(), cancellationToken);
        }

        public DocxApplyResult Apply(
            string handle,
            long expectedVersion,
            string patchText,
            DocxEditOptions? options = null,
            CancellationToken cancellationToken = default)
        {
            long current = GuardVersion(handle, expectedVersion);
            using var input = new MemoryStream(Snapshot(handle, current), writable: false);
            using var patch = new StringReader(patchText);
            using var staged = new MemoryStream();
            DocxApplyResult result = new DocxEditor().Apply(
                input, patch, staged, options ?? new DocxEditOptions(), cancellationToken);
            if (result.Success)
            {
                versionsByHandle[handle].Add(staged.ToArray());
            }

            return result;
        }

        private long GuardVersion(string handle, long expectedVersion)
        {
            List<byte[]> versions = Versions(handle);
            long current = versions.Count - 1;
            if (expectedVersion != current)
            {
                throw new StaleVersionException(expectedVersion, current);
            }

            return current;
        }

        private List<byte[]> Versions(string handle)
        {
            if (!versionsByHandle.TryGetValue(handle, out List<byte[]>? versions))
            {
                throw new InvalidOperationException("Unknown document handle.");
            }

            return versions;
        }
    }

    private static byte[] SeedBytes(string paragraphText)
    {
        using MemoryStream seed = CreateDocx(paragraphText);
        return seed.ToArray();
    }

    private static string[] ReadParagraphs(byte[] docx)
    {
        using var input = new MemoryStream(docx, writable: false);
        DocxReadResult read = new DocxEditor().Read(input);
        Assert.True(read.Success);
        return read.Paragraphs.Select(static paragraph => paragraph.Text).ToArray();
    }

    [Fact]
    public static void ApplyPublishesNewVersionAndKeepsHistoryImmutable()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        Assert.Equal(0, store.CurrentVersion(handle));

        DocxApplyResult result = store.Apply(handle, 0, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        Assert.True(result.Success);
        Assert.Equal(1, store.CurrentVersion(handle));
        Assert.Equal(["Alpha"], ReadParagraphs(store.Snapshot(handle, 0)));
        Assert.Equal(["Beta"], ReadParagraphs(store.Snapshot(handle, 1)));
    }

    [Fact]
    public static void StaleApplyRefusesWithoutDuplication()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        Assert.True(store.Apply(handle, 0, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n").Success);

        StaleVersionException stale = Assert.Throws<StaleVersionException>(() =>
            store.Apply(handle, 0, "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Gamma\nend\n"));
        Assert.Equal(0, stale.ExpectedVersion);
        Assert.Equal(1, stale.ActualVersion);
        Assert.Equal(1, store.CurrentVersion(handle));
        Assert.Equal(["Beta"], ReadParagraphs(store.Snapshot(handle, 1)));

        Assert.True(store.Apply(handle, 1, "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Gamma\nend\n").Success);
        Assert.Equal(2, store.CurrentVersion(handle));
        Assert.Equal(["Beta", "Gamma"], ReadParagraphs(store.Snapshot(handle, 2)));
    }

    [Fact]
    public static void FailedPatchDiscardsStagedOutput()
    {
        var store = new VersionedDocumentStore();
        byte[] original = SeedBytes("Alpha");
        string handle = store.Ingest(original);

        DocxApplyResult result = store.Apply(handle, 0, "docxpatch 1\n\nop replace-text\ntarget M.P9999\nfind Alpha\nwith Beta\nend\n");
        Assert.False(result.Success);
        Assert.Equal(0, store.CurrentVersion(handle));
        Assert.Equal(original, store.Snapshot(handle, 0));
    }

    [Fact]
    public static void CorruptInputFailsWithoutPartialPublication()
    {
        var store = new VersionedDocumentStore();
        byte[] garbage = [0x01, 0x02, 0x03];
        string handle = store.Ingest(garbage);

        DocxApplyResult result = store.Apply(handle, 0, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n");
        Assert.False(result.Success);
        Assert.Equal(0, store.CurrentVersion(handle));
        Assert.Equal(garbage, store.Snapshot(handle, 0));
    }

    [Fact]
    public static void CancellationAbortsWithoutPublication()
    {
        var store = new VersionedDocumentStore();
        byte[] original = SeedBytes("Alpha");
        string handle = store.Ingest(original);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            store.Apply(handle, 0, "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n", cancellationToken: cancelled.Token));
        Assert.Equal(0, store.CurrentVersion(handle));
        Assert.Equal(original, store.Snapshot(handle, 0));
    }

    [Fact]
    public static void HandlesAreOpaqueAndUnknownHandlesFail()
    {
        var store = new VersionedDocumentStore();
        string first = store.Ingest(SeedBytes("Alpha"));
        string second = store.Ingest(SeedBytes("Alpha"));
        Assert.NotEqual(first, second);
        Assert.DoesNotContain(".docx", first, StringComparison.Ordinal);
        Assert.DoesNotContain("/", first, StringComparison.Ordinal);
        Assert.Throws<InvalidOperationException>(() => store.CurrentVersion("missing-handle"));
        Assert.Throws<InvalidOperationException>(() => store.Check("missing-handle", 0, "docxpatch 1\n"));
    }

    [Fact]
    public static void MissingAssetFailsWithoutPublication()
    {
        var store = new VersionedDocumentStore();
        byte[] original = SeedBytes("Alpha");
        string handle = store.Ingest(original);
        var options = new DocxEditOptions
        {
            AssetProvider = new MemoryAssetProvider("other.png", CreatePngBytes(4, 3), null, "other.png")
        };

        DocxApplyResult result = store.Apply(
            handle,
            0,
            "docxpatch 1\n\nop insert-image-after\ntarget M.P0001\nasset missing.png\nalt Chart\nend\n",
            options);
        Assert.False(result.Success);
        Assert.Equal(0, store.CurrentVersion(handle));
        Assert.Equal(original, store.Snapshot(handle, 0));
    }

    [Fact]
    public static void ReconcileUncertainRetryAvoidsDuplicatePublish()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        const string patch = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n";

        // The original request succeeds, but its response is lost in transit.
        store.Apply(handle, 0, patch);

        // The retry reconciles from current state instead of re-applying blindly.
        (bool alreadyApplied, long version) = Reconcile(store, handle, patch, ["Beta"]);
        Assert.True(alreadyApplied);
        Assert.Equal(1, version);
        Assert.Equal(1, store.CurrentVersion(handle));
        Assert.Equal(["Beta"], ReadParagraphs(store.Snapshot(handle, 1)));
    }

    [Fact]
    public static void ReconcileUncertainRetryAppliesWhenAbsent()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        const string patch = "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n";

        // An ambiguous attempt actually failed, and its outcome was lost too.
        store.Apply(handle, 0, "docxpatch 1\n\nop replace-text\ntarget M.P9999\nfind Alpha\nwith Beta\nend\n");

        (bool alreadyApplied, long version) = Reconcile(store, handle, patch, ["Beta"]);
        Assert.False(alreadyApplied);
        Assert.Equal(1, version);
        Assert.Equal(["Beta"], ReadParagraphs(store.Snapshot(handle, 1)));
    }

    private static (bool AlreadyApplied, long Version) Reconcile(
        VersionedDocumentStore store,
        string handle,
        string patchText,
        string[] expectedParagraphs,
        DocxEditOptions? options = null)
    {
        long current = store.CurrentVersion(handle);
        if (ReadParagraphs(store.Snapshot(handle, current)).SequenceEqual(expectedParagraphs))
        {
            return (true, current);
        }

        DocxApplyResult result = store.Apply(handle, current, patchText, options);
        Assert.True(result.Success, string.Join("|", result.Diagnostics.Select(static diagnostic => diagnostic.Code + ":" + diagnostic.Message)));
        return (false, store.CurrentVersion(handle));
    }
}
