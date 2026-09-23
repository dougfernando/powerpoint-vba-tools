Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-ComObject($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Assert-PowerPointClosed {
    if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
        throw 'Feche o PowerPoint antes de continuar. Nenhuma sessao sera encerrada pelo script.'
    }
}

function Get-SourceModel([string]$SourceFolder) {
    $root = (Resolve-Path -LiteralPath $SourceFolder).ProviderPath
    $files = @(Get-ChildItem -LiteralPath $root -File | Where-Object Extension -in '.bas','.cls','.frm' | Sort-Object Name)
    if (!$files.Count) { throw "Nenhum modulo VBA encontrado em $root" }
    $modules = @{}
    foreach ($file in $files) {
        $code = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
        $match = [regex]::Match($code, '(?m)^\s*Attribute VB_Name = "([A-Za-z][A-Za-z0-9_]*)"')
        if (!$match.Success) { throw "VB_Name ausente ou invalido: $($file.Name)" }
        $name = $match.Groups[1].Value
        if ($modules.ContainsKey($name)) { throw "Modulo duplicado: $name" }
        if ($name -in 'modCommandCatalog','frmMacroLauncher') { throw "Modulo reservado: $name" }
        if ($name -ne 'modLoader' -and $code -match '\bThisPresentation\b|\bApplication\.Min\b|\bVBProject\b') {
            throw "Dependencia de desenvolvimento ou API invalida no runtime: $name"
        }
        $modules[$name] = $code
    }
    $catalog = Get-Content (Join-Path $root 'commands.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $catalog = @($catalog)
    if (!$catalog.Count) { throw 'Catalogo vazio.' }
    $ids = @{}
    foreach ($entry in $catalog) {
        foreach ($field in 'id','module','category','label','description') {
            if (!$entry.PSObject.Properties[$field] -or [string]::IsNullOrWhiteSpace($entry.$field)) { throw "Campo $field ausente no catalogo." }
            if ($entry.$field -match '[\r\n]') { throw "Quebra de linha em $field." }
        }
        if ($entry.id -notmatch '^Cmd_[A-Za-z0-9_]+$' -or $entry.module -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { throw "Identificador invalido: $($entry.id)" }
        if ($ids.ContainsKey($entry.id)) { throw "Comando duplicado: $($entry.id)" }
        $ids[$entry.id] = $true
        if (!$modules.ContainsKey($entry.module)) { throw "Modulo inexistente: $($entry.module)" }
        $pattern = '(?im)^\s*Public Sub ' + [regex]::Escape($entry.id) + '\s*\(\s*\)\s*$'
        if ($modules[$entry.module] -notmatch $pattern -or $modules[$entry.module] -match '(?im)^\s*Option Private Module') {
            throw "Comando deve ser Public Sub sem parametros: $($entry.id)"
        }
        if ($entry.PSObject.Properties['scopes']) {
            $scopes = @($entry.scopes)
            if (!$scopes.Count) { throw "Lista de escopos vazia: $($entry.id)" }
            $scopeKeys = @{}
            foreach ($scope in $scopes) {
                if ($scope -notin 'selection','slide','presentation') { throw "Escopo invalido em $($entry.id): $scope" }
                if ($scopeKeys.ContainsKey($scope)) { throw "Escopo duplicado em $($entry.id): $scope" }
                $scopeKeys[$scope] = $true
            }
            if (!$scopeKeys.ContainsKey('slide')) { throw "Comando com escopo deve aceitar slide: $($entry.id)" }
        }
    }
    foreach ($code in $modules.Values) {
        foreach ($match in [regex]::Matches($code, '(?im)^\s*Public Sub (Cmd_\w+)\s*\(')) {
            if (!$ids.ContainsKey($match.Groups[1].Value)) { throw "Comando fora do catalogo: $($match.Groups[1].Value)" }
        }
    }
    $ui = Get-Content (Join-Path $root 'ui\launcher.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $formCode = Get-Content (Join-Path $root 'ui\frmMacroLauncher.code') -Raw -Encoding UTF8
    if ($ui.name -ne 'frmMacroLauncher') { throw 'Nome de formulario invalido.' }
    $controls = @{}
    foreach ($control in $ui.controls) {
        if ($control.type -notin 'Label','ComboBox','ListBox','OptionButton','CommandButton' -or $controls.ContainsKey($control.name)) { throw 'Controle invalido ou duplicado.' }
        $controls[$control.name] = $true
    }
    foreach ($required in 'cboCategory','lstMacro','lblScope','optScopeSelection','optScopeSlide','optScopePresentation','lblDescription','lblResult','cmdRun','cmdClose') {
        if (!$controls.ContainsKey($required)) { throw "Controle ausente: $required" }
    }
    [pscustomobject]@{ Root=$root; Files=$files; Modules=$modules; Catalog=$catalog; UI=$ui; FormCode=$formCode }
}

function ConvertTo-VbaString([string]$Text) { '"' + $Text.Replace('"','""') + '"' }

function New-CatalogCode($Catalog) {
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add('Attribute VB_Name = "modCommandCatalog"')
    $lines.Add('Option Explicit')
    $lines.Add("' Generated from commands.json. Do not edit.")
    $lines.Add('Public Function CommandCatalog() As Collection')
    $lines.Add('    Dim result As New Collection')
    foreach ($entry in $Catalog) {
        $scopeSpec = if ($entry.PSObject.Properties['scopes']) { @($entry.scopes) -join ',' } else { '' }
        $values = @($entry.id,$entry.category,$entry.label,$entry.description,$scopeSpec) | ForEach-Object { ConvertTo-VbaString $_ }
        $lines.Add('    result.Add Array(' + ($values -join ', ') + ')')
    }
    $lines.Add('    Set CommandCatalog = result')
    $lines.Add('End Function')
    $lines.Add('Public Sub DispatchCommand(ByVal commandId As String)')
    $lines.Add('    Select Case commandId')
    foreach ($entry in $Catalog) {
        $lines.Add('        Case ' + (ConvertTo-VbaString $entry.id))
        $lines.Add('            ' + $entry.module + '.' + $entry.id)
    }
    $lines.Add('        Case Else')
    $lines.Add('            Err.Raise 5, "PowerPoint Tools", "Comando desconhecido: " & commandId')
    $lines.Add('    End Select')
    $lines.Add('End Sub')
    $lines -join [Environment]::NewLine
}

function Write-VbaSource([string]$Path, [string]$Code) {
    # VBIDE imports ANSI, while repository sources are UTF-8.
    $encoding = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage,
        [Text.EncoderExceptionFallback]::new(), [Text.DecoderExceptionFallback]::new())
    [IO.File]::WriteAllText($Path, $Code, $encoding)
}

function Set-VbaComponentProperty($Component, [string]$Name, [object]$Value) {
    $properties = $property = $null
    try {
        $properties = $Component.Properties
        $property = $properties.Item($Name)
        # VBIDE.Property.Value is a COM Variant. Invoke its setter directly to
        # avoid PowerShell adapting Value to the type returned by another property.
        [void]$property.GetType().InvokeMember(
            'Value', [Reflection.BindingFlags]::SetProperty, $null,
            $property, [object[]]@($Value))
    } catch {
        throw [InvalidOperationException]::new(
            "Falha na propriedade do formulario '$Name' (tipo $($Value.GetType().Name)): $($_.Exception.Message)",
            $_.Exception)
    } finally {
        Release-ComObject $property
        Release-ComObject $properties
    }
}

function Set-LauncherLayout($Component, $Definition) {
    # Form dimensions belong to VBComponent.Properties, not Designer.
    Set-VbaComponentProperty $Component 'Caption' ([string]$Definition.caption)
    Set-VbaComponentProperty $Component 'Width' ([single]$Definition.width)
    Set-VbaComponentProperty $Component 'Height' ([single]$Definition.height)
    Set-VbaComponentProperty $Component 'StartUpPosition' ([int]1)
}

function Set-LauncherControlLayout($Control, $Definition) {
    $Control.Left = [single]$Definition.left
    $Control.Top = [single]$Definition.top
    $Control.Width = [single]$Definition.width
    $Control.Height = [single]$Definition.height
    if ($Definition.PSObject.Properties['caption']) { $Control.Caption = [string]$Definition.caption }
    if ($Definition.PSObject.Properties['style']) { $Control.Style = [int]$Definition.style }
}

function Import-SourceModel($Project, $Model, [string]$Stage, [string]$AddinFileName = 'DouglasPowerPointTools.ppam') {
    foreach ($file in $Model.Files) {
        $target = Join-Path $Stage $file.Name
        $sourceCode = Get-Content $file.FullName -Raw -Encoding UTF8
        if ($file.BaseName -eq 'modLoader') {
            $embeddedSource = ([string]$Model.Root).Replace('"', '""')
            $embeddedAddin = $AddinFileName.Replace('"', '""')
            $sourceCode = $sourceCode.Replace('{{BUILD_SOURCE_FOLDER}}', $embeddedSource)
            $sourceCode = $sourceCode.Replace('{{BUILD_ADDIN_FILE}}', $embeddedAddin)
        }
        Write-VbaSource $target $sourceCode
        if ($file.Extension -eq '.frm') {
            $frx = [IO.Path]::ChangeExtension($file.FullName,'.frx')
            if (Test-Path -LiteralPath $frx) { Copy-Item -LiteralPath $frx -Destination $Stage }
        }
        $component = $Project.VBComponents.Import($target)
        try {
            $expected = [regex]::Match((Get-Content $file.FullName -Raw -Encoding UTF8), 'Attribute VB_Name = "([^"]+)"').Groups[1].Value
            if ($component.Name -ne $expected) { throw "Nome alterado na importacao: $expected -> $($component.Name)" }
        } finally { Release-ComObject $component }
    }
    $catalogPath = Join-Path $Stage 'modCommandCatalog.bas'
    Write-VbaSource $catalogPath (New-CatalogCode $Model.Catalog)
    $component = $Project.VBComponents.Import($catalogPath)
    Release-ComObject $component
    $form = $null
    $designer = $null
    try {
        $form = $Project.VBComponents.Add(3)
        $form.Name = [string]$Model.UI.name
        Set-LauncherLayout $form $Model.UI
        $designer = $form.Designer
        foreach ($definition in $Model.UI.controls) {
            $control = $designer.Controls.Add("Forms.$($definition.type).1", [string]$definition.name, $true)
            try {
                Set-LauncherControlLayout $control $definition
            } finally { Release-ComObject $control }
        }
        $form.CodeModule.AddFromString([string]$Model.FormCode)
    } finally {
        Release-ComObject $designer
        Release-ComObject $form
    }
    foreach ($reference in $Project.References) {
        try {
            if ($reference.IsBroken) { throw "Referencia VBA ausente: $($reference.Name)" }
        } finally { Release-ComObject $reference }
    }
}

function Publish-File([string]$Candidate, [string]$Destination) {
    if (Test-Path -LiteralPath $Destination) {
        $backup = $Destination + '.' + [Guid]::NewGuid().ToString('N') + '.bak'
        [IO.File]::Replace($Candidate, $Destination, $backup)
        Write-Host "Backup: $backup"
    } else {
        [IO.File]::Move($Candidate, $Destination)
    }
}
