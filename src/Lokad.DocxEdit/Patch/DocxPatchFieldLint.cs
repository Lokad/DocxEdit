namespace Lokad.DocxEdit;

// D10: document-independent patch shape validation. Every rule below is a
// necessary condition of check success: required fields, alternative groups,
// and exclusive groups are all enforced by execution, so lint failures always
// predict check failures while lint success leaves document-dependent guards,
// targets, assets, and shapes to check.
internal static partial class DocxPatchEngine
{
    internal static IReadOnlyList<DocxDiagnostic> ValidatePatchFields(DocxPatch patch)
    {
        var diagnostics = new List<DocxDiagnostic>();
        foreach (DocxPatchOperation operation in patch.Operations)
        {
            if (!OperationsByName.TryGetValue(operation.OperationName, out OperationRegistration? registration))
            {
                continue;
            }

            foreach (OperationFieldDefinition field in registration.Fields.Where(static field => field.Required))
            {
                if (!operation.Fields.TryGetValue(field.Name, out string? value) ||
                    (value.Length == 0 && !field.AllowEmpty))
                {
                    diagnostics.Add(Diagnostic(
                        DocxSeverity.Error,
                        "E4202",
                        "Operation " + Quote(operation.OperationName) + " is missing required field " + Quote(field.Name) + ".",
                        operation));
                }
            }

            string? target = operation.Fields.GetValueOrDefault("target");
            foreach (string[] group in registration.RequireOneOf)
            {
                if (!group.Any(member => IsPresentMember(registration, operation, member)))
                {
                    diagnostics.Add(Diagnostic(
                        DocxSeverity.Error,
                        "E4202",
                        "Operation " + Quote(operation.OperationName) + " requires one of " + JoinQuoted(group) + ".",
                        operation,
                        target));
                }
            }

            foreach (string[] group in registration.ExclusiveGroups)
            {
                if (group.Count(member => IsPresentMember(registration, operation, member)) > 1)
                {
                    diagnostics.Add(Diagnostic(
                        DocxSeverity.Error,
                        "E4205",
                        "Operation " + Quote(operation.OperationName) + " accepts only one of " + JoinQuoted(group) + ".",
                        operation,
                        target));
                }
            }

            if (registration.AcceptedKinds.Length != 0 &&
                operation.Fields.TryGetValue("target", out string? targetValue) &&
                DocxTargetId.TryParse(targetValue, out DocxTargetId parsedTarget) &&
                !registration.AcceptedKinds.Contains(parsedTarget.Kind))
            {
                diagnostics.Add(WrongKindDiagnostic(operation, targetValue!, parsedTarget, DescribeAcceptedKinds(registration.AcceptedKinds)));
            }
        }

        return diagnostics;
    }

    // A member counts as present when its key exists with a value, except
    // that a boolean false counts as absent: execution treats clear false
    // the same as a missing clear field.
    private static bool IsPresentMember(OperationRegistration registration, DocxPatchOperation operation, string member)
    {
        if (!operation.Fields.TryGetValue(member, out string? value) || value.Length == 0)
        {
            return false;
        }

        OperationFieldDefinition? definition = registration.Fields
            .FirstOrDefault(field => string.Equals(field.Name, member, StringComparison.Ordinal));
        if (definition?.Kind == FieldValueKind.Boolean &&
            string.Equals(value, "false", StringComparison.Ordinal))
        {
            return false;
        }

        return true;
    }

    private static string DescribeAcceptedKinds(DocxTargetKind[] kinds)
    {
        return string.Join(" or ", kinds.Select(static kind => DescribeAcceptedKind(kind)));
    }

    private static string DescribeAcceptedKind(DocxTargetKind kind)
    {
        string article = kind == DocxTargetKind.Image ? "an" : "a";
        string example = kind switch
        {
            DocxTargetKind.Paragraph => "M.P0001",
            DocxTargetKind.Table => "M.T0001",
            DocxTargetKind.Row => "M.T0001.R02",
            DocxTargetKind.Cell => "M.T0001.R02.C03",
            DocxTargetKind.MergeGroup => "M.T0001.MG0001",
            DocxTargetKind.Image => "M.I0001",
            DocxTargetKind.Hyperlink => "M.L0001",
            DocxTargetKind.Field => "M.F0001",
            DocxTargetKind.Bookmark => "M.B0001",
            DocxTargetKind.ContentControl => "M.CC0001",
            DocxTargetKind.Section => "M.S0001",
            _ => "M.P0001",
        };
        return article + " " + DescribeTargetKind(kind) + " ID such as " + example;
    }

    private static string JoinQuoted(IReadOnlyList<string> names)
    {
        return string.Join(", ", names.Select(static name => Quote(name)));
    }
}
