<#
  Registers the VisioSizeTools COM add-in for the current user (no admin rights required).
  Usage:  .\Register.ps1 [-Configuration Release|Debug]
  Restart Visio afterwards.
#>
param([ValidateSet('Release', 'Debug')][string]$Configuration = 'Release')

$ErrorActionPreference = 'Stop'
$clsid  = '{4A11EACB-6B73-4915-AD54-D92ECF41F340}'
$progId = 'VisioSizeTools.Connect'
$class  = 'VisioSizeTools.Connect'
$dll    = Join-Path $PSScriptRoot "..\src\VisioSizeTools\bin\$Configuration\net48\VisioSizeTools.dll"

if (-not (Test-Path $dll)) { throw "Build output not found: $dll  (run: dotnet build VisioSizeTools.sln -c $Configuration)" }

$asmName  = [System.Reflection.AssemblyName]::GetAssemblyName($dll).FullName
# Must be an unescaped file URI (as regasm writes it); "%20" for spaces breaks the CLR loader.
$codeBase = 'file:///' + ((Resolve-Path $dll).Path -replace '\\', '/')

function Set-Key($path, [hashtable]$values) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    foreach ($k in $values.Keys) {
        if ($k -eq '(default)') { Set-Item -Path $path -Value $values[$k] }
        else { New-ItemProperty -Path $path -Name $k -Value $values[$k] -PropertyType $(if ($values[$k] -is [int]) { 'DWord' } else { 'String' }) -Force | Out-Null }
    }
}

# COM class registration (HKCU\Software\Classes is shared by 32- and 64-bit Office)
$root = "HKCU:\Software\Classes"
Set-Key "$root\CLSID\$clsid"                 @{ '(default)' = $class }
Set-Key "$root\CLSID\$clsid\ProgId"          @{ '(default)' = $progId }
Set-Key "$root\CLSID\$clsid\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" @{}
Set-Key "$root\CLSID\$clsid\InprocServer32"  @{
    '(default)'      = 'mscoree.dll'
    'ThreadingModel' = 'Both'
    'Class'          = $class
    'Assembly'       = $asmName
    'RuntimeVersion' = 'v4.0.30319'
    'CodeBase'       = $codeBase
}
Set-Key "$root\$progId"       @{ '(default)' = $class }
Set-Key "$root\$progId\CLSID" @{ '(default)' = $clsid }

# Tell Visio to load it at startup
Set-Key "HKCU:\Software\Microsoft\Visio\Addins\$progId" @{
    'FriendlyName' = 'Visio Size Tools'
    'Description'  = 'Size to Tallest / Shortest / Both Larger / Both Smaller on Home > Arrange.'
    'LoadBehavior' = 3
}

Write-Host "Registered $progId -> $dll"
Write-Host 'Restart Visio to load the add-in.'
