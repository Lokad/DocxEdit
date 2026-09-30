using System.Xml.Linq;

namespace Lokad.DocxEdit;

// C02 and C03 shared identity contract. Work in progress. Placement
// enumeration is the single owner for discovery, snapshot, mutation,
// and reporting. Public ordinal IDs are views, not identity.
// A patch local handle names one object for the duration of a patch.
// Input handles bind pre patch objects. Created handles bind objects
// made by the patch. Deleted objects are historical and must not
// resolve to survivors. Do not infer identity twice from neighbors.

internal sealed record PatchObjectHandle(
    DocxTargetKind Kind,
    char Story,
    int StoryPart,
    int Ordinal,
    bool IsCreated,
    Guid InstanceId);

internal sealed record OperationTargetResolution(
    PatchObjectHandle Handle,
    XElement Element,
    string? SnapshotId);
