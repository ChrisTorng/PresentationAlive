# PresentationAlive
Make church presentation easy.

# How to build
Visual Studio 2022 required and recommended.

## How to build in Visual Studio Code
* Open "Developer Command Prompt for VS 2022"
* `cd` to `PresentationAlive\src`
* Run `dotnet build` for restore NuGet library. It will fail, that's ok
* Run `where msbuild` to get the MSBuild.exe path to Visual Studio\2022
* Update `tasks.json`'s `command` path
* Follow the Visual Studio 2022 steps

## How to build in Visual Studio 2022
* Open `regedit`
* Look for all 3 `Guid`s in `PresentationAlive.csproj` under `HKEY_CLASSES_ROOT\TypeLib`
* Update `VersionMajor` and `VersionMinor` by the version under each `Guid`. Hex value (like `b`) should be converted to Decimal (like `11`)
* F5 to run