# .NET 8.0 Upgrade Report — MarsAddIn

## Summary

The upgrade of **MarsAddIn.csproj** from .NET Framework 4.8 to .NET 8.0 was **reverted**. The project is a **VSTO (Visual Studio Tools for Office) Excel Add-in**, which is fundamentally incompatible with .NET 8.0. The VSTO runtime (`Microsoft.Office.Tools`, `Microsoft.Office.Tools.Excel`, `Microsoft.Office.Tools.Common`, etc.) is only available for .NET Framework.

## Why the upgrade was reverted

- VSTO add-ins depend on `Microsoft.Office.Tools` and the VSTO runtime, which have **no .NET 8.0 equivalent**.
- After the initial upgrade converted the project to SDK-style and retargeted to `net8.0-windows`, the build produced **hundreds of errors** due to missing Office interop and VSTO assemblies.
- The project was reverted to its original .NET Framework 4.8 state on the `main` branch.

## What was attempted

- Project was converted to SDK-style format targeting `net8.0-windows`
- Newtonsoft.Json was updated from 13.0.1 to 13.0.4
- VSTO assembly references were removed (as they are not available as NuGet packages for .NET 8.0)
- Build failed with errors in `MarsRibbon.cs`, `MarsRibbon.Designer.cs`, `ThisAddIn.cs`, `ThisAddIn.Designer.cs`, and multiple form files

## Future modernization options

- **Stay on .NET Framework 4.8**: VSTO add-ins work correctly on .NET Framework 4.8. NuGet packages and code patterns can still be modernized within this framework.
- **Migrate to Office Web Add-in (Office JS API)**: Microsoft's recommended modern approach using HTML/CSS/JavaScript. Requires a full rewrite but works cross-platform.
- **Migrate to Excel-DNA**: Open-source framework for building Excel add-ins on .NET 8.0 with ribbon and COM interop support. Requires restructuring but preserves C# codebase.
- **COM Add-in with .NET 8.0**: Use COM interop directly, but loses VSTO designer support and is more complex.

## All commits (on upgrade branch, now reverted)

| Commit ID | Description                                                        |
|:----------|:-------------------------------------------------------------------|
| 2df31dfa  | Commit upgrade plan                                                |
| 06abf08f  | Update MarsAddIn.csproj: Newtonsoft.Json and references            |
| 40211870  | Modernize project: migrate to SDK style and .NET 8                 |
| aa746dfc  | Fix missing Microsoft.Office.Tools references                      |
| d0349ebd  | Add TODO comment on ClickOnce support in HelperClass.cs            |

## Next steps

- Consider updating Newtonsoft.Json from 13.0.1 to 13.0.4 within .NET Framework 4.8
- Evaluate Excel-DNA or Office Web Add-ins if .NET 8.0 migration is a future goal
