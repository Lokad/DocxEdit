# Release procedure

This is a manual release procedure for the first package. No automatic publishing pipeline is required.

## Stable SDK

Use the stable SDK for release builds. The review used stable SDK 10.0.204 alongside the default preview SDK. Select the stable SDK with a temporary global.json pin, then remove the pin immediately after the release build. Do not commit the pin.

Example stable release build from a clean committed tree:

    dotnet build-server shutdown
    dotnet test tests/Lokad.DocxEdit.Tests/Lokad.DocxEdit.Tests.csproj -c Release --nologo -p:TreatWarningsAsErrors=true
    dotnet pack src/Lokad.DocxEdit/Lokad.DocxEdit.csproj -c Release --nologo -p:TreatWarningsAsErrors=true

Run the packaged consumer check once as part of the existing test suite. It packs once to artifacts/nuget and restores a tiny synthetic consumer from that exact local feed with an isolated package cache. It exercises a guarded edit with readback and an image extraction and provider-backed reinsertion. It verifies packaged metadata and version and framework, dependency closure with no runtime dependencies, XML docs, portable symbols, and source commit and path mapping with normalized source paths. Version expectations live in the packaging test constants and the project Version property. Do not add more hard-coded version strings to tests or scripts.

## Artifacts and hashes

The release pack writes nupkg and snupkg files under ignored artifacts/nuget. Retain the tested nupkg and snupkg files with source commit, SDK version, and hashes. Record hashes with:

    Get-FileHash artifacts/nuget/Lokad.DocxEdit.*.nupkg -Algorithm SHA256
    Get-FileHash artifacts/nuget/Lokad.DocxEdit.*.snupkg -Algorithm SHA256

Publish the reviewed artifacts only when release is explicitly requested, then verify install and symbols from the published feed. See official publishing procedure at https://learn.microsoft.com/en-us/nuget/nuget-org/publish-a-package

Do not push, change repository visibility, upload, or release automatically. Report local completion and remaining publication actions distinctly.
