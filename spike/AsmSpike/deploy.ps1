# Builds the spike add-in and registers it with Inventor 2027 (per-user Addins folder).
# Run: pwsh -File deploy.ps1            (build + register)
#      pwsh -File deploy.ps1 -Remove    (unregister)
param([switch]$Remove)

$ErrorActionPreference = 'Stop'
$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$addinDir = Join-Path $env:APPDATA 'Autodesk\Inventor 2027\Addins'
$addinXml = Join-Path $addinDir 'AsmSpike.addin'

if ($Remove) {
    if (Test-Path $addinXml) { Remove-Item $addinXml -Force; "removed $addinXml" } else { "nothing to remove" }
    return
}

dotnet build "$here\AsmSpike.csproj" -c Debug -nologo -v minimal
if ($LASTEXITCODE -ne 0) { throw "build failed" }

$dll = Join-Path $here 'bin\Debug\AsmSpike.dll'
if (-not (Test-Path $dll)) { throw "built dll not found at $dll" }

New-Item -ItemType Directory -Force $addinDir | Out-Null
@"
<?xml version="1.0" encoding="utf-8"?>
<Addin Type="Standard">
  <ClassId>{6B8C2E4A-9D31-4F6E-A0C7-5E2D1B9F3A11}</ClassId>
  <ClientId>{6B8C2E4A-9D31-4F6E-A0C7-5E2D1B9F3A11}</ClientId>
  <DisplayName>AsmSpike (Dynamo ASM coexistence test)</DisplayName>
  <Description>Temporary spike: preloads Dynamo 4.2.1 libG/ASM 232 inside Inventor 2027 and logs the result.</Description>
  <Assembly>$dll</Assembly>
  <LoadOnStartUp>1</LoadOnStartUp>
  <Hidden>0</Hidden>
  <SupportedSoftwareVersionGreaterThan>30..</SupportedSoftwareVersionGreaterThan>
  <UseInventorAssemblyContext>1</UseInventorAssemblyContext>
  <UserUnloadable>1</UserUnloadable>
  <DataVersion>1</DataVersion>
  <UserInterfaceVersion>1</UserInterfaceVersion>
</Addin>
"@ | Set-Content -Path $addinXml -Encoding UTF8
"registered $addinXml -> $dll"
