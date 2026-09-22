<#
Build nativo. Windows PowerShell 5.1, PowerPoint desktop e acesso ao projeto VBA.
Nao altera configuracoes de seguranca nem encerra sessoes existentes.
#>
param(
    [string]$SourceFolder = (Join-Path $PSScriptRoot '..\src'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\build'),
    [string]$OutputFile = 'DouglasPowerPointTools.ppam',
    [switch]$ValidateOnly,
    [switch]$SkipRuntimeValidation
)
. "$PSScriptRoot\lib\common.ps1"
. "$PSScriptRoot\lib\package.ps1"

$model = Get-SourceModel $SourceFolder
$ribbonPath = Join-Path $PSScriptRoot '..\resources\customUI.xml'
[xml]$ribbon = Get-Content $ribbonPath -Raw -Encoding UTF8
if ($OutputFile -ne [IO.Path]::GetFileName($OutputFile) -or [IO.Path]::GetExtension($OutputFile) -ne '.ppam') {
    throw 'OutputFile deve ser apenas um nome de arquivo .ppam.'
}
if ($ValidateOnly) {
    Write-Host "Fontes validas: $($model.Files.Count) modulos, $($model.Catalog.Count) comandos."
    return
}
Assert-PowerPointClosed
$outputRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
[void][IO.Directory]::CreateDirectory($outputRoot)
$stage = Join-Path $outputRoot ('staging-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($stage)
$outputPath = Join-Path $outputRoot $OutputFile
$candidate = Join-Path $stage ('DouglasTools-check-' + [Guid]::NewGuid().ToString('N') + '.ppam')
$ppt = $pres = $project = $null
$step = 'Iniciar PowerPoint'
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    # Do not permit external files to run code while constructing the project.
    $originalSecurity = $ppt.AutomationSecurity
    $ppt.AutomationSecurity = 3
    $pres = $ppt.Presentations.Add(0)
    $slide = $pres.Slides.Add(1, 12)
    Release-ComObject $slide
    $step = 'Criar e reabrir container PPTM'
    $containerPath = Join-Path $stage 'DouglasPowerPointTools.build.pptm'
    $pres.SaveAs($containerPath, 25)
    $pres.Close()
    Release-ComObject $pres
    $pres = $null
    $pres = $ppt.Presentations.Open($containerPath, 0, 0, 0)
    if ($pres.ReadOnly) { throw 'Container PPTM aberto somente para leitura.' }
    $step = 'Acessar projeto VBA'
    $project = $pres.VBProject
    if (!$project -or $project.Protection -ne 0) { throw 'Projeto VBA indisponivel ou protegido. Verifique o acesso ao modelo VBA.' }
    $step = 'Importar fontes e construir formulario'
    Import-SourceModel -Project $project -Model $model -Stage $stage -AddinFileName $OutputFile
    Release-ComObject $project
    $project = $null
    $step = 'Salvar PPTM de desenvolvimento'
    $pres.Save()
    $step = 'Salvar PPAM nativo (formato 30)'
    try {
        $pres.SaveAs($candidate, 30)
        $pres.Close()
        Release-ComObject $pres
        $pres = $null
        Write-Host 'PPAM salvo pelo formato nativo 30.'
    } catch [Runtime.InteropServices.COMException] {
        $nativeSaveError = $_
        Write-Warning "O PowerPoint recusou o formato 30 (HRESULT $($nativeSaveError.Exception.HResult)). Usando conversao Open XML preservando o projeto VBA."
        if (Test-Path -LiteralPath $candidate) { Remove-Item -LiteralPath $candidate -Force }
        # The PPTM was saved immediately before the native attempt. Do not save
        # again because a failed SaveAs may leave ambiguous COM path state.
        $pres.Saved = -1
        $pres.Close()
        Release-ComObject $pres
        $pres = $null
        $step = 'Converter PPTM salvo para PPAM (fallback Open XML)'
        Convert-PptmPackageToPpam $containerPath $candidate
    }
    $step = 'Incorporar e validar RibbonX'
    Add-Ribbon $candidate $ribbonPath
    Test-PpamPackage $candidate
    $ppt.AutomationSecurity = $originalSecurity
    if ($SkipRuntimeValidation) {
        Write-Warning 'Validacao de carregamento ignorada. Instale e teste o launcher manualmente no PowerPoint.'
    } else {
        $step = 'Carregar add-in e verificar formulario/catalogo'
        Test-LoadedTools $ppt $candidate
    }
    $step = 'Encerrar PowerPoint antes de publicar'
    $ppt.Quit()
    Release-ComObject $ppt
    $ppt = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # A previously registered add-in may point to the output path. PowerPoint
    # holds that PPAM open until its process has fully exited.
    $exitDeadline = [DateTime]::UtcNow.AddSeconds(10)
    while (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
        if ([DateTime]::UtcNow -ge $exitDeadline) {
            throw 'A instancia de PowerPoint criada pelo build nao encerrou em 10 segundos. O PPAM de destino continua em uso.'
        }
        Start-Sleep -Milliseconds 250
    }

    $step = 'Publicar PPAM'
    Publish-File $candidate $outputPath
    Write-Host "PPAM gerado e carregamento verificado: $outputPath"
    Write-Host "PPTM e fontes intermediarias: $stage"
    Write-Host 'Conclua os testes visuais e funcionais de docs/VALIDATION.md.'
} catch {
    $message = "Etapa: $step | HRESULT: $($_.Exception.HResult) | $($_.Exception.Message)"
    $logPath = Join-Path $stage 'build-error.log'
    $details = @(
        $message
        $_.InvocationInfo.PositionMessage
        $_.ScriptStackTrace
        $_.Exception.ToString()
    ) -join [Environment]::NewLine
    [IO.File]::WriteAllText($logPath, $details)
    Write-Warning "$message | Log: $logPath"
    throw
} finally {
    Release-ComObject $project
    if ($null -ne $pres) {
        try { $pres.Saved = -1; $pres.Close() } finally { Release-ComObject $pres }
    }
    if ($null -ne $ppt) {
        try { $ppt.Quit() } finally { Release-ComObject $ppt }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
