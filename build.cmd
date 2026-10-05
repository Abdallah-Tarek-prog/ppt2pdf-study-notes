@echo off
rem Builds ppt2pdf.exe with the C# compiler that ships with Windows (.NET Framework 4).
rem Nothing needs to be installed.
setlocal
set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
    echo Could not find the .NET Framework 4 C# compiler.
    exit /b 1
)
"%CSC%" /nologo /target:exe /platform:anycpu /optimize+ /warnaserror+ ^
    /out:"%~dp0ppt2pdf.exe" ^
    /reference:Microsoft.CSharp.dll /reference:System.Core.dll /reference:System.Drawing.dll ^
    "%~dp0src\*.cs"
if errorlevel 1 exit /b 1
echo Built %~dp0ppt2pdf.exe
