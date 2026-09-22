param(
    [Parameter(Mandatory=$true)][string]$PresentationPath,
    [string]$SourceFolder = (Join-Path $PSScriptRoot '..\src')
)
. "$PSScriptRoot\lib\common.ps1"
. "$PSScriptRoot\lib\development.ps1"
$model = Get-SourceModel $SourceFolder
$stage = Join-Path ([IO.Path]::GetTempPath()) ('DouglasTools-import-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($stage)
$ppt = $pres = $project = $null
$backup = Join-Path $stage 'before-import.pptm'
try {
    $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    $pres = Get-ExplicitPresentation $ppt $PresentationPath
    if ($pres.ReadOnly) { throw 'O PPTM esta aberto somente para leitura.' }
    $project = $pres.VBProject
    if ($project.Protection -ne 0) { throw 'Projeto protegido.' }
    $managed = @()
    for ($i = 1; $i -le $project.VBComponents.Count; $i++) {
        $component = $project.VBComponents.Item($i)
        try {
            $count = [Math]::Min(20, $component.CodeModule.CountOfLines)
            $header = if ($count -gt 0) { $component.CodeModule.Lines(1, $count) } else { '' }
            $owned = ($header -match '@ManagedByDouglasTools') -or ($component.Name -in 'modLoader','modCommandCatalog','frmMacroLauncher')
            if ($owned -and $component.Type -ne 100) { $managed += $component.Name }
            elseif ($model.Modules.ContainsKey($component.Name)) { throw "Colisao com modulo nao gerenciado: $($component.Name)" }
        } finally { Release-ComObject $component }
    }
    $pres.SaveCopyAs($backup, 25)
    foreach ($name in $managed) {
        $component = $project.VBComponents.Item($name)
        try { $project.VBComponents.Remove($component) } finally { Release-ComObject $component }
    }
    Import-SourceModel $project $model $stage
    Write-Host "Importado em: $($pres.FullName). Revise e compile antes de salvar."
    Write-Host "Backup completo anterior: $backup"
} catch {
    Write-Warning "Se a importacao iniciou, restaure o PPTM anterior a partir de: $backup"
    throw
} finally {
    Release-ComObject $project
    Release-ComObject $pres
    Release-ComObject $ppt
    # Never close the user's PowerPoint or save changes automatically.
}
