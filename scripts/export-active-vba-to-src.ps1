param(
    [Parameter(Mandatory=$true)][string]$PresentationPath,
    [string]$SourceFolder = (Join-Path $PSScriptRoot '..\src')
)
. "$PSScriptRoot\lib\common.ps1"
. "$PSScriptRoot\lib\development.ps1"
$root = (Resolve-Path -LiteralPath $SourceFolder).ProviderPath
$backupRoot = Join-Path $PSScriptRoot '..\build'
[void][IO.Directory]::CreateDirectory($backupRoot)
$stage = Join-Path $backupRoot ('export-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($stage)
$ppt = $pres = $project = $null
try {
    $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $pres = Get-ExplicitPresentation $ppt $PresentationPath
    $project = $pres.VBProject
    if ($project.Protection -ne 0) { throw 'Projeto protegido.' }
    $exportFolder = Join-Path $stage 'exported'
    Export-Components $project $exportFolder
    # Only update existing managed source modules. Exported UI/catalog stay in staging
    # because their source of truth is JSON and the form event code.
    $updates = @()
    foreach ($file in Get-ChildItem -LiteralPath $exportFolder -Filter '*.bas' -File) {
        if ($file.BaseName -in 'modLoader','modCommandCatalog') { continue }
        $destination = Join-Path $root $file.Name
        if (!(Test-Path -LiteralPath $destination)) { continue }
        $existing = Get-Content $destination -Raw -Encoding UTF8
        $incoming = Get-Content $file.FullName -Raw -Encoding UTF8
        if ($existing -notmatch '@ManagedByDouglasTools' -or $incoming -notmatch '@ManagedByDouglasTools') {
            throw "Modulo nao gerenciado: $($file.Name)"
        }
        $updates += $file
    }
    if (!$updates.Count) { throw 'Nenhum modulo gerenciado correspondente encontrado.' }
    $backup = Join-Path $stage 'source-backup'
    Copy-Item -LiteralPath $root -Destination $backup -Recurse
    try {
        foreach ($file in $updates) {
            Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $root $file.Name) -Force
        }
        $null = Get-SourceModel $root
    } catch {
        foreach ($file in $updates) {
            Copy-Item -LiteralPath (Join-Path $backup $file.Name) -Destination (Join-Path $root $file.Name) -Force
        }
        throw
    }
    Write-Host "Fontes atualizadas: $($updates.Count). Backup: $backup"
    Write-Host "Exportacao completa para revisao (inclui formulario e catalogo): $exportFolder"
} finally {
    Release-ComObject $project
    Release-ComObject $pres
    Release-ComObject $ppt
}
