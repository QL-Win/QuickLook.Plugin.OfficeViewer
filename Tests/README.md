# Legacy PowerPoint integration smoke test

Requires Windows, LibreOffice in a standard Program Files location, and a C#
compiler targeting .NET Framework 4.6.2 or later. No Syncfusion key or QuickLook
installation is required for this converter test.

From a Visual Studio Developer Command Prompt, at the repository root:

```bat
mkdir bin\SmokeTest
csc /nologo /out:bin\SmokeTest\SmokeTest.exe LegacyPowerPointConverter.cs Tests\LegacyPowerPointSmokeTest.cs
bin\SmokeTest\SmokeTest.exe bin\SmokeTest\Cache "C:\path\to\presentation.ppt"
```

Use a fresh cache directory to exercise actual conversion. More input paths can
be supplied. The test checks concurrent conversion requests, cache reuse without
rewriting the PDF, pre-cancellation, and preservation of source contents. It does
not exercise WPF preview lifecycle or assess slide rendering fidelity.

Manual checks with QuickLook: preview a `.ppt`, switch documents or close the
window during conversion, reopen after completion, and verify that the page
counter uses the original filename. Existing `.pptx`, Word and Excel previews
should continue to use their current renderer.
