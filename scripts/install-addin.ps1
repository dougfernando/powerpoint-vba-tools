param(
    [string]$PpamPath = (Join-Path $PSScriptRoot '..\build\DouglasPowerPointTools.ppam'),
    [switch]$SkipRuntimeValidation
)
. "$PSScriptRoot\lib\common.ps1"
. "$PSScriptRoot\lib\package.ps1"
$source = (Resolve-Path -LiteralPath $PpamPath).ProviderPath
Test-PpamPackage $source
Assert-PowerPointClosed
$folder = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Microsoft\AddIns\DouglasPowerPointTools'
[void][IO.Directory]::CreateDirectory($folder)
$target = Join-Path $folder 'DouglasPowerPointTools.ppam'
if ($source -eq $target) { throw 'Use o PPAM da pasta build como origem.' }
$token = [Guid]::NewGuid().ToString('N')
$candidate = Join-Path $folder ("DouglasTools-check-$token.ppam")
$backup = Join-Path $folder ("DouglasPowerPointTools.$token.bak")
$hadFile = Test-Path -LiteralPath $target
$published = $false
$changedRegistration = $false
$previousRegistered = 0
$previousLoaded = 0
$ppt = $addin = $null
Copy-Item -LiteralPath $source -Destination $candidate
if ($hadFile) { Copy-Item -LiteralPath $target -Destination $backup }
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    if (!$SkipRuntimeValidation) { Test-LoadedTools $ppt $candidate }
    foreach ($item in $ppt.AddIns) {
        if ([IO.Path]::GetFileName($item.FullName) -eq 'DouglasPowerPointTools.ppam') {
            if ($item.FullName -ne $target) { throw "Outro add-in com o mesmo nome esta registrado: $($item.FullName). Remova esse registro antes de instalar." }
            $addin = $item
            $previousRegistered = $addin.Registered
            $previousLoaded = $addin.Loaded
            break
        }
        Release-ComObject $item
    }
    if ($null -ne $addin) {
        $changedRegistration = $true
        $addin.Loaded = 0
        $addin.Registered = 0
        $ppt.AddIns.Remove($addin.Name)
        Release-ComObject $addin
        $addin = $null
    }
    Publish-File $candidate $target
    $published = $true
    $changedRegistration = $true
    $addin = $ppt.AddIns.Add($target)
    $addin.Loaded = -1
    if (!$SkipRuntimeValidation) {
        $result = Invoke-PowerPointMacro $ppt 'DouglasPowerPointTools.ppam!modMacroLauncher.ToolsHealthCheck'
        if ($result -ne 'DouglasPowerPointTools:OK') { throw 'Verificacao da instalacao falhou.' }
    } else {
        Write-Warning 'Health check ignorado. Confirme manualmente a aba PowerPoint Tools e o launcher.'
    }
    $addin.Registered = -1
    Write-Host "Instalado: $target"
    if ($hadFile) { Write-Host "Versao anterior: $backup" }
} catch {
    $failure = $_
    try {
        if ($changedRegistration -and $null -ne $addin) {
            $addin.Loaded = 0
            $addin.Registered = 0
            $ppt.AddIns.Remove($addin.Name)
            Release-ComObject $addin
            $addin = $null
        }
        if ($published) {
            if ($hadFile) { Copy-Item -LiteralPath $backup -Destination $target -Force }
            else { Move-Item -LiteralPath $target -Destination ($target + ".$token.failed") }
        }
        if ($changedRegistration -and $hadFile -and $null -ne $ppt) {
            $addin = $ppt.AddIns.Add($target)
            $addin.Registered = $previousRegistered
            $addin.Loaded = $previousLoaded
        }
    } catch { Write-Warning "Falha na restauracao: $($_.Exception.Message). Backup: $backup" }
    throw $failure
} finally {
    Release-ComObject $addin
    if ($null -ne $ppt) { try { $ppt.Quit() } finally { Release-ComObject $ppt } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
