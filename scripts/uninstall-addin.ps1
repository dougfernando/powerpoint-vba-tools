param()
. "$PSScriptRoot\lib\common.ps1"
Assert-PowerPointClosed
$folder = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Microsoft\AddIns\DouglasPowerPointTools'
$target = Join-Path $folder 'DouglasPowerPointTools.ppam'
$ppt = $addin = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    foreach ($item in $ppt.AddIns) {
        if ($item.FullName -eq $target) { $addin = $item; break }
        Release-ComObject $item
    }
    if ($null -ne $addin) {
        $addin.Loaded = 0
        $addin.Registered = 0
        $ppt.AddIns.Remove($addin.Name)
    }
} finally {
    Release-ComObject $addin
    if ($null -ne $ppt) { try { $ppt.Quit() } finally { Release-ComObject $ppt } }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
if (Test-Path -LiteralPath $target) {
    $backup = $target + '.' + [Guid]::NewGuid().ToString('N') + '.uninstalled'
    Move-Item -LiteralPath $target -Destination $backup
    Write-Host "Registro removido. Arquivo preservado em: $backup"
} else { Write-Host 'Add-in nao instalado neste caminho.' }
