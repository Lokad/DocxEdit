namespace Lokad.DocxEdit.OfficeTests;

public sealed class OfficeFactAttribute : FactAttribute
{
    public OfficeFactAttribute()
    {
        if (!string.Equals(Environment.GetEnvironmentVariable("DOCXEDIT_ENABLE_OFFICE_TESTS"), "1", StringComparison.Ordinal))
        {
            Skip = "Office Word tests are disabled.";
        }
    }
}
