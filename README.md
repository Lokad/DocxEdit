# DocxEdit

DocxEdit is a stream-first net10.0 library for inspecting and editing docx files with deterministic patch files. It has no runtime package dependencies beyond the .NET platform libraries.

Install the library with:

    dotnet add package Lokad.DocxEdit --version 0.1.0

Requires net10.0. Installing this package does not install the docxedit command. For the development CLI, clone the source and run via dotnet run, see source CLI guidance at https://github.com/Lokad/DocxEdit/blob/master/docs/cli.md

## Library example

The example stages output in memory, checks Success with diagnostics, and publishes only on success. Caller-owned streams stay open by default. Wrap file creation in try catch for publication and input output errors. The CLI preserves any pre-existing destination file byte for byte on failed apply. Library callers get similar staging safety as below: staging keeps a failed Apply from reaching the destination, but File.Create followed by CopyToAsync is not atomic and an input output failure during publication can leave a partial file. For CLI equivalent file safety, publish through a temporary file in the same directory and move it over the destination only after the copy succeeds.

    using Lokad.DocxEdit;

    await using Stream input = File.OpenRead("report.docx");
    using var patch = File.OpenText("edits.docxpatch");
    await using var staged = new MemoryStream();
    DocxApplyResult result = new DocxEditor().Apply(input, patch, staged);
    if (!result.Success)
    {
        foreach (DocxDiagnostic diagnostic in result.Diagnostics)
        {
            Console.WriteLine(diagnostic.Code + ": " + diagnostic.Message);
        }

        return;
    }

    staged.Position = 0;
    await using Stream output = File.Create("report.edited.docx");
    await staged.CopyToAsync(output);

Input and output streams are left open by default. Set LeaveInputOpen or LeaveOutputOpen to false on the relevant options when the editor should dispose them.

## Essential guidance

Assets: image edits use external assets through AssetProvider. The library never reads asset files directly. Extract bytes with ExtractMedia, edit bytes with external tools, and reinsert with replace-image or insert-image-after using a provider that returns the new bytes. PNG and JPEG interchange updates the media content type. An unshared media part keeps its existing part path, which may retain the old extension, while the exported file name follows the actual content type. Treat part paths as opaque and do not rename package parts to match this prose.

IDs: target IDs enumerate physical document order at inspection time from read output. They are stable for the same document bytes and scanner version, but not across modifications. Within one patch, explicit IDs bind to the input snapshot, semantic selectors resolve live against current content, and created objects are addressed through patch-local aliases with whole target @name. Reports carry Id with Coordinate input or operation-time and FinalId live matching a fresh read, staying absent for deleted objects. A deleted image or link never acquires its surviving containing paragraph identity. Check and apply agree on Success, diagnostics, and per-operation reports under the same input and options, including affected targets, created IDs, and previews. GeneratedRevisionIds remain empty in check by design and are populated only by apply. Re-read output for final coordinates.

Tracking: track-changes modes are off, suggest, and require. Only some operations generate revision markup with author, timestamp, and revision IDs in reports. Annotation operations such as comments stay permitted under Require. Complex shapes warn or fail with explicit diagnostics. See patch guidance at https://github.com/Lokad/DocxEdit/blob/master/docs/patch-format.md

Limits: there are no layout, rendering, pagination, or Word fidelity guarantees. Word open and save checks are bounded compatibility only, not proof for arbitrary documents. macOS execution, deployed storage behavior, live agent usability, and NuGet publishing rights are not established by local tests.

## Documentation

Source guidance lives at:

https://github.com/Lokad/DocxEdit/blob/master/docs/cli.md
https://github.com/Lokad/DocxEdit/blob/master/docs/patch-format.md
https://github.com/Lokad/DocxEdit/blob/master/docs/diagnostics.md
https://github.com/Lokad/DocxEdit/blob/master/docs/validation.md
https://github.com/Lokad/DocxEdit/blob/master/docs/status.md

For product integrations, the same guidance is available from the library through DocxHelp Catalog, and CLI text output is reusable through DocxTextRenderer.

## Source build and CLI

The package itself contains only the net10.0 library with XML docs, README, changelog, license, and icon. For source development:

    dotnet build Lokad.DocxEdit.slnx
    dotnet test Lokad.DocxEdit.slnx
    dotnet pack src/Lokad.DocxEdit/Lokad.DocxEdit.csproj -c Release

Release packs generate both nupkg and snupkg files under ignored artifacts/nuget. Non-Release packs are rejected unless AllowNonReleasePackage true is supplied for troubleshooting. Run the CLI from source with dotnet run, as shown in source CLI guidance. No tool package is published.

CI runs the Release library and package-consumer tests on Windows and Ubuntu, treats warnings as errors, and retains a test log for each platform. Word open and save tests require an explicit opt-in on Windows with Word installed.
