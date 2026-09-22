Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

function Read-ZipText($Zip, [string]$Name) {
    $entry = $Zip.GetEntry($Name)
    if (!$entry) { throw "Parte ausente: $Name" }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Set-ZipText($Zip, [string]$Name, [string]$Text) {
    $old = $null
    if ($Zip.Mode -ne [IO.Compression.ZipArchiveMode]::Create) { $old = $Zip.GetEntry($Name) }
    if ($old) { $old.Delete() }
    $entry = $Zip.CreateEntry($Name)
    $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Convert-PptmPackageToPpam([string]$SourcePath, [string]$DestinationPath) {
    $source = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    if ([IO.Path]::GetExtension($source) -ne '.pptm') { throw 'A origem do fallback deve ser um PPTM.' }
    if ([IO.Path]::GetExtension($DestinationPath) -ne '.ppam') { throw 'O destino do fallback deve ser um PPAM.' }

    Copy-Item -LiteralPath $source -Destination $DestinationPath
    $zip = [IO.Compression.ZipFile]::Open($DestinationPath, [IO.Compression.ZipArchiveMode]::Update)
    try {
        [xml]$contentTypes = Read-ZipText $zip '[Content_Types].xml'
        $mainParts = @($contentTypes.DocumentElement.ChildNodes | Where-Object {
            $_.GetAttribute('PartName') -eq '/ppt/presentation.xml'
        })
        if ($mainParts.Count -ne 1) { throw 'Parte principal da apresentacao ausente ou duplicada.' }

        $pptmType = 'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml'
        $ppamType = 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml'
        if ($mainParts[0].GetAttribute('ContentType') -ne $pptmType) {
            throw 'O container nao possui o content type esperado de PPTM.'
        }
        $mainParts[0].SetAttribute('ContentType', $ppamType)
        Set-ZipText $zip '[Content_Types].xml' $contentTypes.OuterXml
    } catch {
        $failure = $_
        $zip.Dispose()
        $zip = $null
        if (Test-Path -LiteralPath $DestinationPath) { Remove-Item -LiteralPath $DestinationPath -Force }
        throw $failure
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
    }
}

function Add-Ribbon([string]$Path, [string]$RibbonPath) {
    $ribbon = Get-Content -LiteralPath $RibbonPath -Raw -Encoding UTF8
    [xml]$null = $ribbon
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        [xml]$rels = Read-ZipText $zip '_rels/.rels'
        $ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        $type = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
        foreach ($existing in @($rels.DocumentElement.ChildNodes)) {
            if ($existing.GetAttribute('Type') -eq $type) { [void]$rels.DocumentElement.RemoveChild($existing) }
        }
        $rel = $rels.CreateElement('Relationship', $ns)
        $rel.SetAttribute('Id', 'rIdDouglasTools' + [Guid]::NewGuid().ToString('N'))
        $rel.SetAttribute('Type', $type)
        $rel.SetAttribute('Target', 'customUI/customUI.xml')
        [void]$rels.DocumentElement.AppendChild($rel)
        [xml]$ct = Read-ZipText $zip '[Content_Types].xml'
        $part = '/customUI/customUI.xml'
        foreach ($existing in @($ct.DocumentElement.ChildNodes)) {
            if ($existing.GetAttribute('PartName') -eq $part) { [void]$ct.DocumentElement.RemoveChild($existing) }
        }
        $override = $ct.CreateElement('Override', 'http://schemas.openxmlformats.org/package/2006/content-types')
        $override.SetAttribute('PartName', $part)
        $override.SetAttribute('ContentType', 'application/xml')
        [void]$ct.DocumentElement.AppendChild($override)
        Set-ZipText $zip 'customUI/customUI.xml' $ribbon
        Set-ZipText $zip '_rels/.rels' $rels.OuterXml
        Set-ZipText $zip '[Content_Types].xml' $ct.OuterXml
    } finally { $zip.Dispose() }
}

function Test-PpamPackage([string]$Path) {
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $names = @{}
        foreach ($entry in $zip.Entries) {
            if ($names.ContainsKey($entry.FullName)) { throw "Parte duplicada: $($entry.FullName)" }
            $names[$entry.FullName] = $true
        }
        [xml]$ct = Read-ZipText $zip '[Content_Types].xml'
        $main = @($ct.DocumentElement.ChildNodes | Where-Object { $_.GetAttribute('PartName') -eq '/ppt/presentation.xml' })
        if ($main.Count -ne 1 -or $main[0].GetAttribute('ContentType') -ne 'application/vnd.ms-powerpoint.addin.macroEnabled.main+xml') { throw 'Formato nativo PPAM ausente.' }
        $vba = $zip.GetEntry('ppt/vbaProject.bin')
        if (!$vba -or !$vba.Length) { throw 'Projeto VBA ausente ou vazio.' }
        [xml]$ribbon = Read-ZipText $zip 'customUI/customUI.xml'
        if ($ribbon.DocumentElement.NamespaceURI -ne 'http://schemas.microsoft.com/office/2006/01/customui') { throw 'Namespace Ribbon invalido.' }
        $buttons = @($ribbon.SelectNodes("//*[local-name()='button' and @onAction='RibbonShowMacroLauncher']"))
        if ($buttons.Count -ne 1) { throw 'Callback do launcher ausente ou duplicado.' }
        $reloadButtons = @($ribbon.SelectNodes("//*[local-name()='button' and @onAction='RibbonReloadMacros']"))
        if ($reloadButtons.Count -ne 1) { throw 'Callback de recarga ausente ou duplicado.' }
        $dfsTabs = @($ribbon.SelectNodes("//*[local-name()='tab' and @label='DFS Tools']"))
        if ($dfsTabs.Count -ne 1) { throw 'Aba DFS Tools ausente ou duplicada.' }
        $reloadButtons = @($ribbon.SelectNodes("//*[local-name()='button' and @onAction='RibbonReloadMacros']"))
        if ($reloadButtons.Count -ne 1) { throw 'Callback de recarga ausente ou duplicado.' }
        $dfsTabs = @($ribbon.SelectNodes("//*[local-name()='tab' and @label='DFS Tools']"))
        if ($dfsTabs.Count -ne 1) { throw 'Aba DFS Tools ausente ou duplicada.' }
        [xml]$rels = Read-ZipText $zip '_rels/.rels'
        $uiRels = @($rels.DocumentElement.ChildNodes | Where-Object {
            $_.GetAttribute('Type') -eq 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility' -and
            $_.GetAttribute('Target') -eq 'customUI/customUI.xml'
        })
        if ($uiRels.Count -ne 1) { throw 'Relacionamento Ribbon ausente ou duplicado.' }
        $uiTypes = @($ct.DocumentElement.ChildNodes | Where-Object {
            $_.GetAttribute('PartName') -eq '/customUI/customUI.xml' -and $_.GetAttribute('ContentType') -eq 'application/xml'
        })
        if ($uiTypes.Count -ne 1) { throw 'Content type Ribbon ausente.' }
    } finally { $zip.Dispose() }
}

function Invoke-PowerPointMacro($Application, [string]$MacroName) {
    # PowerPoint exposes Run(MacroName, safeArrayOfParams). VBA treats the
    # second argument as optional, but the PowerShell COM binder does not.
    $noArguments = [object[]]@()
    return $Application.Run($MacroName, $noArguments)
}

function Test-LoadedTools($Application, [string]$Path) {
    $addin = $null
    try {
        $addin = $Application.AddIns.Add($Path)
        $addin.Loaded = -1
        $result = Invoke-PowerPointMacro $Application ([IO.Path]::GetFileName($Path) + '!modMacroLauncher.ToolsHealthCheck')
        if ($result -ne 'DouglasPowerPointTools:OK') { throw 'Health check do add-in falhou.' }
    } finally {
        if ($null -ne $addin) {
            $addin.Loaded = 0
            $addin.Registered = 0
            $Application.AddIns.Remove($addin.Name)
            Release-ComObject $addin
        }
    }
}
