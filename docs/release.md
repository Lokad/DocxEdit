# Release procedure

This is a manual release procedure for the first package. No automatic publishing pipeline is required.

## Stable SDK

Use the stable SDK for release builds. Select the stable SDK with a temporary global.json pin, then remove the pin immediately after the release build. Do not commit the pin. The packaging test propagates the effective repository SDK version into the isolated consumer directory with its own global.json pin and asserts both report the same version, so a consumer under a system temporary directory cannot silently use a different SDK.

Example stable release build from a clean committed tree (the test pack is the release pack; do not pack again after it):

    dotnet build-server shutdown
    dotnet test tests/Lokad.DocxEdit.Tests/Lokad.DocxEdit.Tests.csproj -c Release --nologo -p:TreatWarningsAsErrors=true --filter PackagingTests

The packaging test packs once to artifacts/nuget and restores a tiny synthetic consumer from that exact local feed with an isolated package cache. It exercises a guarded edit with readback and an image extraction and provider-backed reinsertion. It parses the nuspec XML and the portable PDB with System.Reflection.Metadata and asserts the expected source commit, SourceLink mapping, normalized document paths, version, framework, and dependency closure with no runtime dependencies, plus XML docs and portable symbols. The project Version property is the single owner of the release version; the test derives its expectations from the project file. Do not add hard-coded version strings to tests or scripts.

## Artifacts and hashes

The test pack writes nupkg and snupkg files under ignored artifacts/nuget. Retain and hash those exact tested files without repacking. Every commit changes package bytes through the nuspec repository commit, the assembly informational version, and PDB SourceLink, so hashes are only valid for the recorded commit and must be retaken after any change, including test-only commits. Record the source revision and SDK versions alongside the hashes:

    git rev-parse HEAD
    dotnet --version
    Get-FileHash artifacts/nuget/Lokad.DocxEdit.*.nupkg -Algorithm SHA256
    Get-FileHash artifacts/nuget/Lokad.DocxEdit.*.snupkg -Algorithm SHA256

Publish the reviewed artifacts only when release is explicitly requested, then verify install and symbols from the published feed. See official publishing procedure at https://learn.microsoft.com/en-us/nuget/nuget-org/publish-a-package

Do not push, change repository visibility, upload, or release automatically. Report local completion and remaining publication actions distinctly.
