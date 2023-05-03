# PresentationAlive
Make church presentation easy.

# How to build
Please note that Visual Studio 2022 (Community up) with .NET Desktop workload is required.
While it's possible to use VS Code for development, we highly recommend using Visual Studio 2022.

## How to build in Visual Studio Code
1. Open "Developer Command Prompt for VS 2022"
2. Change the directory to "PresentationAlive\src" using the `cd` command.
3. Run `dotnet restore` to restore the NuGet libraries.
4. Run `where msbuild` to get the path of the `MSBuild.exe` file under `Microsoft Visual Studio\2022` folder.
5. Update the command path in the `tasks.json` file, change path divider `\` to `/`.
6. Follow the steps specific to Visual Studio 2022, or do the following manual steps.

Manual steps:
1. Navigate to the file src\PowerPointLib\PowerPointLib.csproj
   Currently thay are:
   +----------------------------------------+-------------------------------------+-------------+
   | {0002E157-0000-0000-C000-000000000046} | VBIDE                               | 5.3         |
   +----------------------------------------+-------------------------------------+-------------+
   | {2DF8D04C-5BFA-101B-BDE5-00AA0044DE52} | Microsoft.Office.Core               | 2.7         |
   +----------------------------------------+-------------------------------------+-------------+
   | {91493440-5A91-11CF-8700-00AA0060263B} | Microsoft.Office.Interop.PowerPoint | 2.b => 2.11 |
   +----------------------------------------+-------------------------------------+-------------+
2. Open regedit and navigate to `HKEY_CLASSES_ROOT\TypeLib`
3. Search for the GUID values obtained in step 1
4. Update the `VersionMajor` and `VersionMinor` values in the `PowerPointLib.csproj` file based on the version information obtained in step 1
   If the version is in hex format (e.g., 2.b), convert it to decimal (e.g., 2.11)

## How to build in Visual Studio 2022
1. Open sln in Visual Studio 2022
2. Navigate to the Dependencies\COM node under the PowerPointLib project
3. If there are any exclamation marks, right-click on the COM node and select "Add COM Reference..."
4. Fix the exclamation marks by selecting the following options:
   * Interop.Microsoft.Office.Interop.PowerPoint
     => Microsoft Office 15.0 Object Library: 2.7
   * Microsoft.Vbe.Interop
     => Microsoft PowerPoint 15.0 Object Library: 2.11
   * Office
     => Microsoft Visual Basic for Applications Extensibility: 5.3