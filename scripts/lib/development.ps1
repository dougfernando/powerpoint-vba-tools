function Get-ExplicitPresentation($Application, [string]$PresentationPath) {
    $path = (Resolve-Path -LiteralPath $PresentationPath).ProviderPath
    if ([IO.Path]::GetExtension($path) -ne '.pptm') { throw 'O alvo de desenvolvimento deve ser um PPTM salvo.' }
    foreach ($presentation in $Application.Presentations) {
        if ($presentation.FullName -eq $path) { return $presentation }
        Release-ComObject $presentation
    }
    throw "Abra o PPTM indicado no PowerPoint: $path"
}

function Export-Components($Project, [string]$Folder) {
    [void][IO.Directory]::CreateDirectory($Folder)
    foreach ($component in $Project.VBComponents) {
        try {
            $extension = switch ($component.Type) { 1 { '.bas' } 2 { '.cls' } 3 { '.frm' } default { $null } }
            if (!$extension) { continue }
            $path = Join-Path $Folder ($component.Name + $extension)
            $component.Export($path)
            # VBIDE exports ANSI. Normalize text for source control, preserving FRX bytes.
            $encoding = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage)
            $text = [IO.File]::ReadAllText($path, $encoding)
            [IO.File]::WriteAllText($path, $text, [Text.UTF8Encoding]::new($false))
        } finally { Release-ComObject $component }
    }
}
