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

    private sealed record CommitRecord(long Version, DocxApplyResult Result, string PatchText);

    private sealed class VersionedDocumentStore
    {
        private readonly object gate = new();
        private readonly Dictionary<string, List<byte[]>> versionsByHandle = new(StringComparer.Ordinal);
        private readonly Dictionary<string, Dictionary<string, CommitRecord>> commitsByHandle = new(StringComparer.Ordinal);

        public string Ingest(byte[] docx)
        {
            string handle = Guid.NewGuid().ToString("N");
            lock (gate)
            {
                versionsByHandle[handle] = [(byte[])docx.Clone()];
            }

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
            CancellationToken cancellationToken = default,
            string? requestId = null,
            Action? onStaged = null)
        {
            byte[] basis;
            lock (gate)
            {
                if (requestId is not null
                    && commitsByHandle.TryGetValue(handle, out Dictionary<string, CommitRecord>? prior)
                    && prior.TryGetValue(requestId, out CommitRecord? record))
                {
                    if (!string.Equals(record.PatchText, patchText, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException($"Request {requestId} was already committed with different patch text.");
                    }

                    return record.Result;
                }

                long current = GuardVersion(handle, expectedVersion);
                basis = Snapshot(handle, current);
            }

            using var input = new MemoryStream(basis, writable: false);
            using var patch = new StringReader(patchText);
            using var staged = new MemoryStream();
            DocxApplyResult result = new DocxEditor().Apply(
                input, patch, staged, options ?? new DocxEditOptions(), cancellationToken);
            onStaged?.Invoke();
            if (!result.Success)
            {
                return result;
            }

            lock (gate)
            {
                long current = GuardVersion(handle, expectedVersion);
                if (current != expectedVersion)
                {
                    throw new StaleVersionException(expectedVersion, current);
                }

                versionsByHandle[handle].Add(staged.ToArray());
                long published = versionsByHandle[handle].Count - 1;
                DocxApplyResult published_result = result;
                if (requestId is not null)
                {
                    if (!commitsByHandle.TryGetValue(handle, out Dictionary<string, CommitRecord>? perHandle))
                    {
                        perHandle = new Dictionary<string, CommitRecord>(StringComparer.Ordinal);
                        commitsByHandle[handle] = perHandle;
                    }

                    perHandle[requestId] = new CommitRecord(published, published_result, patchText);
                }
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
            lock (gate)
            {
                if (!versionsByHandle.TryGetValue(handle, out List<byte[]>? versions))
                {
                    throw new InvalidOperationException("Unknown document handle.");
                }

                return versions;
            }
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
    public static void PublicationIsAtomicCompareAndSwap()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        StaleVersionException stale = Assert.Throws<StaleVersionException>(() =>
            store.Apply(
                handle,
                0,
                "docxpatch 1\n\nop replace-text\ntarget M.P0001\nfind Alpha\nwith Beta\nend\n",
                requestId: "loser",
                onStaged: () => Assert.True(store.Apply(
                    handle,
                    0,
                    "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Gamma\nend\n").Success)));
        Assert.Equal(0, stale.ExpectedVersion);
        Assert.Equal(1, stale.ActualVersion);
        Assert.Equal(1, store.CurrentVersion(handle));
        Assert.Equal(["Alpha", "Gamma"], ReadParagraphs(store.Snapshot(handle, 1)));
    }

    [Fact]
    public static void RepeatingRequestReturnsCommittedOutcomeWithoutDuplication()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        const string patch = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Gamma\nend\n";
        DocxApplyResult first = store.Apply(handle, 0, patch, requestId: "req-1");
        Assert.True(first.Success);
        Assert.Equal(1, store.CurrentVersion(handle));
        DocxApplyResult replay = store.Apply(handle, 0, patch, requestId: "req-1");
        Assert.True(replay.Success);
        Assert.Equal(1, store.CurrentVersion(handle));
        Assert.Equal(["Alpha", "Gamma"], ReadParagraphs(store.Snapshot(handle, 1)));
    }

    [Fact]
    public static void AbsentAttemptRetryAppliesOnce()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        const string bad = "docxpatch 1\n\nop replace-text\ntarget M.P9999\nfind Alpha\nwith Beta\nend\n";
        Assert.False(store.Apply(handle, 0, bad, requestId: "req-absent").Success);
        Assert.Equal(0, store.CurrentVersion(handle));
        const string good = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Gamma\nend\n";
        Assert.True(store.Apply(handle, 0, good, requestId: "req-absent").Success);
        Assert.Equal(1, store.CurrentVersion(handle));
        Assert.Equal(["Alpha", "Gamma"], ReadParagraphs(store.Snapshot(handle, 1)));
    }

    [Fact]
    public static void ConflictingRequestWithStaleVersionRefuses()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        Assert.True(store.Apply(handle, 0, "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Gamma\nend\n", requestId: "req-a").Success);
        StaleVersionException stale = Assert.Throws<StaleVersionException>(() =>
            store.Apply(handle, 0, "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Delta\nend\n", requestId: "req-b"));
        Assert.Equal(0, stale.ExpectedVersion);
        Assert.Equal(1, stale.ActualVersion);
        Assert.Equal(["Alpha", "Gamma"], ReadParagraphs(store.Snapshot(handle, 1)));
    }

    [Fact]
    public static void ConflictingPatchTextUnderSameRequestRefuses()
    {
        var store = new VersionedDocumentStore();
        string handle = store.Ingest(SeedBytes("Alpha"));
        const string first = "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Gamma\nend\n";
        Assert.True(store.Apply(handle, 0, first, requestId: "req-same").Success);
        Assert.Throws<InvalidOperationException>(() =>
            store.Apply(handle, 0, "docxpatch 1\n\nop insert-after\ntarget M.P0001\ntext Delta\nend\n", requestId: "req-same"));
        Assert.Equal(1, store.CurrentVersion(handle));
        Assert.Equal(["Alpha", "Gamma"], ReadParagraphs(store.Snapshot(handle, 1)));
    }

    [Fact]
    public static void ReferenceDeclaresTestedLibraryContract()
    {
        System.Reflection.AssemblyName name = typeof(DocxEditor).Assembly.GetName();
        Assert.Equal("Lokad.DocxEdit", name.Name);
        Assert.NotNull(name.Version);
    }


}
