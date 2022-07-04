# PresentationAlive
Make church presentation easy.

# How to build
Visual Studio 2022 required and recommended.

## How to build in Visual Studio Code
* Open "Developer Command Prompt for VS 2022"
* `cd` to `PresentationAlive\src`
* Run `dotnet restore` for restore NuGet library
* Run `where msbuild` to get the MSBuild.exe path used by Visual Studio 2022
* Update `tasks.json`'s `command` path
* Follow the Visual Studio 2022 steps

## How to build in Visual Studio 2022
* Open `regedit`
* Look for all 3 `Guid`s in `src\PowerPointLib\PowerPointLib.csproj` under `HKEY_CLASSES_ROOT\TypeLib`
* Update `VersionMajor` and `VersionMinor` by the version under each `Guid`. Hex value (like `b`) should be converted to Decimal (like `11`)
* F5 to run