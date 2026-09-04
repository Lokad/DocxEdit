namespace Lokad.DocxEdit.Tests;

public static class OperationRegistryTests
{
    [Fact]
    public static void RegistryCatalogAndEngineCoverTheSameOperations()
    {
        string[] registry = DocxPatchEngine.AllOperations.Select(registration => registration.Name).ToArray();
        string[] catalog = DocxHelp.Catalog.PatchOperations.Select(operation => operation.Name).ToArray();

        Assert.Equal(registry, catalog);

        foreach (string name in registry)
        {
            Assert.True(DocxPatchEngine.IsKnownOperation(name));
            Assert.True(
                DocxPatchEngine.TryGetPatchOperationHandler(name, out PatchOperationHandler? handler) &&
                handler is not null);

            DocxPatch patch = DocxPatchParser.Parse($"docxpatch 1\n\nop {name}\nend\n");
            Assert.DoesNotContain(patch.Diagnostics, diagnostic => diagnostic.Code == "E2010");
        }

        Assert.False(DocxPatchEngine.IsKnownOperation("no-such-operation"));
    }

    [Fact]
    public static void RegistryFieldSchemasAreInternallyConsistent()
    {
        foreach (OperationRegistration registration in DocxPatchEngine.AllOperations)
        {
            var allowed = new HashSet<string>(registration.AllowedFields, StringComparer.Ordinal);
            foreach (string booleanField in registration.BooleanFields)
            {
                Assert.Contains(booleanField, allowed);
            }

            foreach (string integerField in registration.IntegerFields)
            {
                Assert.Contains(integerField, allowed);
            }

            var documented = new HashSet<string>(
                registration.Catalog.RequiredFields.Concat(registration.Catalog.OptionalFields),
                StringComparer.Ordinal);
            bool proseCover = registration.Catalog.RequiredFields
                .Any(field => field.Contains(" plus ", StringComparison.Ordinal) || field.Contains(" or ", StringComparison.Ordinal));
            foreach (string field in allowed)
            {
                Assert.True(
                    documented.Contains(field) || proseCover,
                    $"Operation '{registration.Name}' allows field '{field}' with no catalog trace.");
            }

            Assert.Equal(registration.Name, registration.Catalog.Name);
        }
    }

    [Fact]
    public static void TrackedChangeClassificationAgreesBetweenEngineAndCatalog()
    {
        foreach (DocxPatchOperationInfo operation in DocxHelp.Catalog.PatchOperations)
        {
            Assert.Equal(
                operation.GeneratesTrackedChanges,
                DocxPatchEngine.SupportsTrackedChangeOutput(operation.Name));
        }
    }

    [Fact]
    public static void FieldDirtyMarkingMatchesDocumentedOperations()
    {
        Assert.False(DocxPatchEngine.MarksFieldsDirtyAfterEdit("set-field-result"));
        Assert.False(DocxPatchEngine.MarksFieldsDirtyAfterEdit("refresh-field-result"));
        Assert.True(DocxPatchEngine.MarksFieldsDirtyAfterEdit("replace-text"));
        Assert.True(DocxPatchEngine.MarksFieldsDirtyAfterEdit("append-row"));
    }
}
