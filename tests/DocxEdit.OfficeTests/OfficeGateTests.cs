namespace DocxEdit.OfficeTests;

public static class OfficeGateTests
{
    [Fact]
    public static void OfficeAutomationIsOptIn()
    {
        if (!string.Equals(Environment.GetEnvironmentVariable("DOCXEDIT_ENABLE_OFFICE_TESTS"), "1", StringComparison.Ordinal))
        {
            return;
        }

        if (!OperatingSystem.IsWindows())
        {
            throw new InvalidOperationException("Office integration tests require Windows.");
        }

        Type? wordApplicationType = Type.GetTypeFromProgID("Word.Application");
        if (wordApplicationType is null)
        {
            throw new InvalidOperationException("Microsoft Word is not installed or is not available through COM.");
        }
    }
}

