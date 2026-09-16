# Registers the Debug build of DynamoInventor with Inventor 2027 for the current user.
#
#   pwsh -File scripts\deploy-dev-addin.ps1            build + register
#   pwsh -File scripts\deploy-dev-addin.ps1 -NoBuild   register only
#   pwsh -File scripts\deploy-dev-addin.ps1 -Remove    unregister
#
# The manifest written to %APPDATA%\Autodesk\Inventor 2027\Addins points at the DLL in bin\Debug
# with an absolute path, so rebuilding is enough to pick up changes (restart Inventor).
# The Dynamo runtime is found via DYNAMO_INVENTOR_RUNTIME, a DynamoCore folder next to the DLL,
# or %ProgramFiles%\Dynamo\Dynamo Core\4.x (see DynamoRuntime.RootFolder).
param(
    [switch]$Remove,
    [switch]$NoBuild,
    [string]$Configuration = 'Debug',
    [string]$InventorVersion = '2027',
    [string]$DynamoRuntime = 'C:\AFAutomations\Tools\DynamoCoreRuntime\4.2.1'
)

$ErrorActionPreference = 'Stop'
$repo     = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$sln      = Join-Path $repo 'src\DynamoInventor.slnx'
$dll      = Join-Path $repo "bin\$Configuration\DynamoInventor.dll"
$addinDir = Join-Path $env:APPDATA "Autodesk\Inventor $InventorVersion\Addins"
$addinXml = Join-Path $addinDir 'Autodesk.DynamoInventor.Inventor.dev.addin'

if ($Remove) {
    if (Test-Path $addinXml) { Remove-Item $addinXml -Force; "removed $addinXml" } else { "nothing to remove" }
    return
}

if (-not $NoBuild) {
    dotnet build $sln -c $Configuration -nologo -v minimal
    if ($LASTEXITCODE -ne 0) { throw "build failed" }
}
if (-not (Test-Path $dll)) { throw "built dll not found at $dll" }

# Tell DynamoInventor.App where the Dynamo runtime is (dev machines: the extracted DynamoCoreRuntime zip).
if ($DynamoRuntime) {
    Set-Content -Path (Join-Path (Split-Path $dll) 'dynamo-runtime.txt') -Value $DynamoRuntime -Encoding ASCII -NoNewline
    "runtime pointer -> $DynamoRuntime"
}

# Take the shipped manifest and swap the relative Assembly path for the absolute dev path.
$template = Get-Content (Join-Path $repo 'src\DynamoInventor\Autodesk.DynamoInventor.Inventor.addin') -Raw
$manifest = $template -replace '<Assembly>[^<]*</Assembly>', "<Assembly>$dll</Assembly>" `
                      -replace '<DisplayName>[^<]*</DisplayName>', '<DisplayName>Dynamo for Inventor (DEV)</DisplayName>'

New-Item -ItemType Directory -Force $addinDir | Out-Null
Set-Content -Path $addinXml -Value $manifest -Encoding UTF8
"registered $addinXml -> $dll"
