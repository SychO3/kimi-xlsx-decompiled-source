# KimiXlsx decompiled source

Source: thvroyal/kimi-skills `skills/kimi-xlsx/scripts/KimiXlsx`
SHA256: a3c7331ac248effdc269641b22b55fa91b9b3fee6f208f222877e3a4e74d8aff

Extraction/decompile workflow:
- Identified as .NET 8 single-file linux-x64 ELF bundle.
- Extracted bundle header at offset `0x496cd55`; bundle version 6.0; 179 embedded files.
- Decompiled `KimiXlsx.dll` with `ilspycmd 9.1.0.7988`.
- Minimal compile fixes applied: removed compiler-only `RefSafetyRules` assembly marker and restored tuple element names lost by decompiler in two places.
- Project references NuGet packages instead of embedded DLL copies: ClosedXML 0.105.0, DocumentFormat.OpenXml 3.1.1.

Verified on Linux with .NET SDK 8.0.421: `dotnet build -c Release` succeeded.
