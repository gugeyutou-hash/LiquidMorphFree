param(
    [string]$Dist = $PSScriptRoot,
    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'
$clsid = '{7647A779-F679-48FF-8A74-A853D0233B01}'
$progId = 'LiquidMorph.Free.Desktop'
$addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\AddIns\$progId"
$clsidKey = "HKCU:\Software\Classes\CLSID\$clsid"
$progIdKey = "HKCU:\Software\Classes\$progId"

if ($Unregister) {
    Remove-Item $addinKey -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $clsidKey -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $progIdKey -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host 'LiquidMorph add-in unregistered.'
    return
}

$dll = Join-Path $Dist 'LiquidMorph.Desktop.dll'
if (-not (Test-Path $dll)) { throw "Missing $dll" }

$assembly = [Reflection.AssemblyName]::GetAssemblyName($dll)
$fullName = $assembly.FullName

New-Item -Path "$clsidKey\InprocServer32" -Force | Out-Null
Set-ItemProperty -Path $clsidKey -Name '(default)' -Value $progId
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name '(default)' -Value "$env:WINDIR\System32\mscoree.dll"
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'ThreadingModel' -Value 'Both'
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'Class' -Value 'LiquidMorph.Addin.Connect'
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'Assembly' -Value $fullName
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path "$clsidKey\InprocServer32" -Name 'CodeBase' -Value ([uri]$dll).AbsoluteUri

New-Item -Path "$progIdKey\CLSID" -Force | Out-Null
Set-ItemProperty -Path "$progIdKey\CLSID" -Name '(default)' -Value $clsid
Remove-ItemProperty -Path $progIdKey -Name 'CLSID' -ErrorAction SilentlyContinue

New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'LiquidMorph Free'
Set-ItemProperty -Path $addinKey -Name 'Description' -Value 'Free open-source liquid glass effect for PowerPoint'
Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord

Write-Host "Registered $progId"
Write-Host "  DLL: $dll"
Write-Host 'Open PowerPoint and look for the LiquidMorph Free tab.'



