# Windows Build Instructions

## Prerequisites
- Visual Studio 2019 or later (or Visual Studio Build Tools)
- .NET Framework 4.8 SDK
- Microsoft Word installed on the build machine

## Build Steps

### 1. Clean Previous Builds
Before building, clean any cached files:

```powershell
# Delete bin and obj folders
Remove-Item -Recurse -Force bin, obj -ErrorAction SilentlyContinue

# Clean Visual Studio cache
dotnet clean
```

### 2. Restore NuGet Packages
```powershell
# Restore packages using NuGet
nuget restore BuildingBlocksManager.sln

# Or using dotnet CLI
dotnet restore
```

### 3. Build the Project
```powershell
# Build using MSBuild
msbuild BuildingBlocksManager.sln /p:Configuration=Release /p:Platform="Any CPU"

# Or using dotnet CLI
dotnet build --configuration Release
```

### 4. Alternative: Use Traditional .NET Framework Project
If the SDK-style project has issues, use the traditional project:
```powershell
# Build the traditional Framework project
msbuild BuildingBlocksManager.Framework.csproj /p:Configuration=Release
```

## Troubleshooting

### Assembly Attribute Errors
If you see duplicate assembly attribute errors:
1. Delete the entire `obj` folder
2. Clean the solution
3. Rebuild from scratch

### Package Restore Issues
If NuGet packages fail to restore:
1. Check internet connection
2. Clear NuGet cache: `nuget locals all -clear`
3. Restore packages manually: `nuget restore`

### Word COM Interop Issues
- Ensure Microsoft Word is installed
- Run Visual Studio as Administrator if needed
- Check that Office Interop assemblies are available

## Output Location
- **Debug**: `bin\Debug\BuildingBlocksManager.exe`
- **Release**: `bin\Release\BuildingBlocksManager.exe`

## Testing
1. Create a test folder with some `AT_*.docx` files
2. Create a test Word template (.dotm file)
3. Run the application and test basic functionality
4. Check logs in `%USERPROFILE%\AppData\Local\BuildingBlocksManager\Logs\`

## Package Versions
- Microsoft.Office.Interop.Word: 15.0.4797.1003
- System.Text.Json: 8.0.5 (latest secure version)

## Notes
- The application targets .NET Framework 4.8
- COM Interop requires Word to be installed
- All file operations are logged for debugging
- Template backups are created automatically