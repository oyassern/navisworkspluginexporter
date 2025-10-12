@echo off
echo Building NavisExcelExporter Plugin...

REM Restore NuGet packages
echo Restoring NuGet packages...
nuget restore packages.config -PackagesDirectory packages

REM Build the project
echo Building project...
msbuild NavisExcelExporter.csproj /p:Configuration=Release /p:Platform="x64"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful!
    echo.
    echo Output files:
    echo - bin\Release\NavisExcelExporter.dll
    echo - NavisExcelExporter.addin
    echo.
    echo To deploy:
    echo 1. Copy both files to: %%APPDATA%%\Autodesk Navisworks Manage 2023\Plugins\
    echo 2. Restart Navisworks Manage
    echo 3. Find the plugin under Tools ^> External Tools
) else (
    echo.
    echo Build failed! Please check the error messages above.
)

pause
