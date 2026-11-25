$ErrorActionPreference='Stop'
$ms=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/MSBuild.exe'
foreach($project in 'src/host/LiquidMorph.Host.csproj','src/addin/LiquidMorph.Addin.csproj'){
    & $ms (Join-Path $PSScriptRoot $project) /nologo /p:Configuration=Release
    if($LASTEXITCODE){exit $LASTEXITCODE}
}
