# Offline tests: no Office automation, no installation and no user files modified.
. "$PSScriptRoot\..\scripts\lib\common.ps1"
. "$PSScriptRoot\..\scripts\lib\package.ps1"
$repo = Split-Path $PSScriptRoot -Parent
$temp = Join-Path ([IO.Path]::GetTempPath()) ('DouglasTools-tests-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
$script:passed = 0

function Assert-True([bool]$Condition, [string]$Message) {
    if (!$Condition) { throw "FAIL: $Message" }
    $script:passed++
}
function Assert-Throws([scriptblock]$Action, [string]$Pattern) {
    $caught = $false
    try { & $Action | Out-Null } catch {
        if ($_.Exception.Message -notmatch $Pattern) { throw "Unexpected error: $($_.Exception.Message); expected $Pattern" }
        $caught = $true
    }
    Assert-True $caught "Expected error: $Pattern"
}
function Write-TestText([string]$Path,[string]$Text) {
    [IO.File]::WriteAllText($Path,$Text,[Text.UTF8Encoding]::new($false))
}
function New-Fixture {
    $folder = Join-Path $temp ([Guid]::NewGuid().ToString('N'))
    Copy-Item -LiteralPath (Join-Path $repo 'src') -Destination $folder -Recurse
    $folder
}

foreach ($file in Get-ChildItem (Join-Path $repo 'scripts'),$PSScriptRoot -Filter '*.ps1' -Recurse) {
    $tokens = $null; $errors = $null
    $null = [Management.Automation.Language.Parser]::ParseFile($file.FullName,[ref]$tokens,[ref]$errors)
    Assert-True ($errors.Count -eq 0) "Syntax: $($file.Name): $errors"
}
$model = Get-SourceModel (Join-Path $repo 'src')
Assert-True ($model.Catalog.Count -eq 15) '15 shipped commands'
Assert-True (@($model.Catalog | Where-Object id -eq 'Cmd_Text_RemoveManualLineBreaks').Count -eq 1) 'Manual line-break command in catalog'
Assert-True ($model.Modules.ContainsKey('modNotifications')) 'Notification channel included in source model'
Assert-True ($model.Modules['modNotifications'] -match 'Public Sub Notify\(ByVal message As String\)') 'Notification publisher is available'
Assert-True ($model.Modules['modNotifications'] -match 'Public Function ConsumeNotification\(\) As String') 'Notification consumer is available'
Assert-True ($model.Modules['modNotifications'] -match 'RESULT_DISPLAY_TIME_MS As Long = 1000') 'Result is cleared after one second'
Assert-True ($model.Modules['modNotifications'] -match 'AddressOf NotificationTimerCallback') 'Result timer uses asynchronous callback'
Assert-True (@($model.UI.controls | Where-Object name -eq 'lblResult').Count -eq 1) 'Launcher result label'
Assert-True (@($model.UI.controls | Where-Object { $_.name -eq 'lstMacro' -and $_.type -eq 'ListBox' }).Count -eq 1) 'Visible macro list'
Assert-True ($model.Modules['modMacroLauncher'] -match 'CStr\(item\(2\)\)\s*&\s*" \("\s*&\s*CStr\(item\(1\)\)') 'Command label precedes category'
foreach ($moduleCode in $model.Modules.Values) {
    Assert-True ($moduleCode -notmatch '\bMsgBox\b') 'Runtime module does not use MsgBox'
}
Assert-True ($model.FormCode -notmatch '\bMsgBox\b') 'Launcher form does not use MsgBox'
foreach ($entry in $model.Catalog) {
    $moduleCode = $model.Modules[$entry.module]
    $procedurePattern = '(?ms)^Public Sub ' + [regex]::Escape($entry.id) + '\(\)(.*?)(?=^End Sub)'
    $procedure = [regex]::Match($moduleCode, $procedurePattern)
    Assert-True $procedure.Success "Command procedure found: $($entry.id)"
    Assert-True ($procedure.Value -match '\bNotify\b') "Command publishes a result: $($entry.id)"
}
Assert-True ($model.Modules.ContainsKey('modLoader')) 'Runtime loader included in source model'
Assert-True ($model.Modules['modLoader'] -match '\{\{BUILD_SOURCE_FOLDER\}\}') 'Loader source path placeholder present'
Assert-True ($model.Modules['modLoader'] -match '\{\{BUILD_ADDIN_FILE\}\}') 'Loader add-in name placeholder present'
Assert-True ($model.Modules['modLoader'] -notmatch '\bThisPresentation\b') 'Loader avoids unsupported ThisPresentation global'
Assert-True ($model.Modules['modLoader'] -match 'Application\.VBE\.VBProjects') 'Loader resolves host through VBE projects'
Assert-True ($model.Modules['modLoader'] -notmatch 'Presentations\(BUILD_ADDIN_FILE\)') 'Loader avoids hidden add-in lookup through Presentations'
[xml]$ribbonSource = Get-Content (Join-Path $repo 'resources\customUI.xml') -Raw -Encoding UTF8
Assert-True (@($ribbonSource.SelectNodes("//*[local-name()='tab' and @label='DFS Tools']")).Count -eq 1) 'DFS Tools tab'
Assert-True (@($ribbonSource.SelectNodes("//*[local-name()='button' and @onAction='RibbonReloadMacros']")).Count -eq 1) 'Reload Ribbon button'
Assert-True ($model.Modules.ContainsKey('modLoader')) 'Runtime loader included in source model'
Assert-True ($model.Modules['modLoader'] -match '\{\{BUILD_SOURCE_FOLDER\}\}') 'Loader source path placeholder present'
[xml]$ribbonSource = Get-Content (Join-Path $repo 'resources\customUI.xml') -Raw -Encoding UTF8
Assert-True (@($ribbonSource.SelectNodes("//*[local-name()='tab' and @label='DFS Tools']")).Count -eq 1) 'DFS Tools tab'
Assert-True (@($ribbonSource.SelectNodes("//*[local-name()='button' and @onAction='RibbonReloadMacros']")).Count -eq 1) 'Reload Ribbon button'
if (-not ('ToolsTest.PowerPointApplication' -as [type])) {
    Add-Type -TypeDefinition @"
namespace ToolsTest {
    public class PowerPointApplication {
        public object Run(string macroName, object[] args) {
            if (args == null || args.Length != 0) throw new System.Exception("Expected empty SafeArray.");
            return macroName;
        }
    }
}
"@
}
$runProbe = New-Object ToolsTest.PowerPointApplication
Assert-True ((Invoke-PowerPointMacro $runProbe 'Addin.ppam!Module.Check') -eq 'Addin.ppam!Module.Check') 'PowerPoint Run receives explicit empty SafeArray'
# Model the VBIDE boundary: component -> Properties.Item(name) -> Variant Value.
# There are deliberately no Width/Height/Caption properties on the component.
# This is an offline contract test, not a substitute for real Office automation.
if (-not ('ToolsTest.VbaComponent' -as [type])) {
    Add-Type -TypeDefinition @"
namespace ToolsTest {
    public class VbaProperty {
        public object Value { get; set; }
    }
    public class VbaProperties {
        private System.Collections.Generic.Dictionary<string, VbaProperty> items =
            new System.Collections.Generic.Dictionary<string, VbaProperty>();
        public VbaProperties() {
            foreach (string name in new [] { "Caption", "Width", "Height", "StartUpPosition" })
                items.Add(name, new VbaProperty());
        }
        public VbaProperty Item(string name) { return items[name]; }
    }
    public class VbaComponent {
        public VbaProperties Properties { get; private set; }
        public VbaComponent() { Properties = new VbaProperties(); }
    }
}
"@
}
$component = New-Object ToolsTest.VbaComponent
Set-LauncherLayout $component $model.UI
$caption = $component.Properties.Item('Caption').Value
$width = $component.Properties.Item('Width').Value
$height = $component.Properties.Item('Height').Value
$startUp = $component.Properties.Item('StartUpPosition').Value
Assert-True ($caption -is [string] -and $caption -eq 'DFS Tools') 'Form caption through VBIDE is String'
Assert-True ($width -is [single] -and $width -eq 440) 'Form width through VBIDE is Single'
Assert-True ($height -is [single] -and $height -eq 418) 'Form height through VBIDE is Single'
Assert-True ($model.FormCode -match 'cmdRun\.Default\s*=\s*True') 'Enter runs selected command'
Assert-True ($model.FormCode -match 'cmdClose\.Cancel\s*=\s*True') 'Escape closes launcher'
Assert-True ($model.FormCode -match 'cboCategory\.TabIndex\s*=\s*0') 'Keyboard navigation starts at category'
Assert-True ($model.FormCode -match 'lstMacro\.TabIndex\s*=\s*1') 'Keyboard navigation continues to command'
Assert-True ($model.FormCode -match 'UserForm_Activate[\s\S]*lstMacro\.SetFocus') 'Initial focus moves to macro list'
Assert-True ($model.FormCode -match 'Case vbKeyC[\s\S]*cboCategory\.SetFocus') 'Alt+C focuses category explicitly'
Assert-True ($model.FormCode -match 'Case vbKeyM[\s\S]*lstMacro\.SetFocus') 'Alt+M focuses macro list explicitly'
Assert-True ($model.FormCode -match 'ScheduleNotificationClear') 'Launcher schedules result cleanup'
Assert-True ($startUp -is [int] -and $startUp -eq 1) 'Form startup position through VBIDE is Int32'
Assert-Throws { Set-VbaComponentProperty $component 'MissingProperty' 1 } 'MissingProperty'
foreach ($definition in $model.UI.controls) {
    $control = [pscustomobject]@{ Left=$null; Top=$null; Width=$null; Height=$null; Caption=$null; Style=$null }
    Set-LauncherControlLayout $control $definition
    Assert-True ($control.Left -is [single] -and $control.Top -is [single] -and $control.Width -is [single] -and $control.Height -is [single]) "Control geometry types: $($definition.name)"
    if ($definition.PSObject.Properties['caption']) {
        Assert-True ($control.Caption -is [string] -and $control.Caption -eq $definition.caption) "Control caption: $($definition.name)"
    }
    if ($definition.PSObject.Properties['style']) {
        Assert-True ($control.Style -is [int] -and $control.Style -eq 2) "ComboBox style: $($definition.name)"
    }
}
$generated = New-CatalogCode $model.Catalog
Assert-True ($generated -notmatch 'Application.Run|VBProject|ThisPresentation') 'Direct dispatch only'
foreach ($entry in $model.Catalog) {
    Assert-True ($generated.Contains($entry.module + '.' + $entry.id)) "Dispatch target $($entry.id)"
}
Assert-True ((ConvertTo-VbaString 'a"b') -eq '"a""b"') 'VBA string quoting'
Assert-Throws { Get-SourceModel (Join-Path $temp 'absent') } 'does not exist|nao existe|n.o existe|Cannot find|localizar|encontr'
$fixture = New-Fixture
Copy-Item (Join-Path $fixture 'modConfig.bas') (Join-Path $fixture 'duplicate.bas')
Assert-Throws { Get-SourceModel $fixture } 'Modulo duplicado'
$fixture = New-Fixture
$entries = Get-Content (Join-Path $fixture 'commands.json') -Raw | ConvertFrom-Json
$entries = @($entries)
Write-TestText (Join-Path $fixture 'commands.json') (ConvertTo-Json -InputObject @($entries + $entries[0]) -Depth 5)
Assert-Throws { Get-SourceModel $fixture } 'Comando duplicado'
$fixture = New-Fixture
$entries[0].module = 'modMissing'
Write-TestText (Join-Path $fixture 'commands.json') (ConvertTo-Json -InputObject $entries -Depth 5)
Assert-Throws { Get-SourceModel $fixture } 'Modulo inexistente'
$fixture = New-Fixture
$path = Join-Path $fixture 'modCommands_Text.bas'
Write-TestText $path ((Get-Content $path -Raw) -replace 'Public Sub Cmd_Text_Swap\(\)', 'Public Sub Cmd_Text_Swap(ByVal x As Long)')
Assert-Throws { Get-SourceModel $fixture } 'sem parametros'
$fixture = New-Fixture
$path = Join-Path $fixture 'modCommands_Text.bas'
Write-TestText $path ((Get-Content $path -Raw) + [Environment]::NewLine + 'Public Sub Cmd_Test_New()')
Assert-Throws { Get-SourceModel $fixture } 'fora do catalogo'
$fixture = New-Fixture
$path = Join-Path $fixture 'modConfig.bas'
Write-TestText $path ((Get-Content $path -Raw) + [Environment]::NewLine + 'Set p = ThisPresentation')
Assert-Throws { Get-SourceModel $fixture } 'API invalida'
$fixture = New-Fixture
$path = Join-Path $fixture 'ui\launcher.json'
$ui = Get-Content $path -Raw | ConvertFrom-Json
$ui.controls[0].name = 'cmdRun'
Write-TestText $path ($ui | ConvertTo-Json -Depth 5)
Assert-Throws { Get-SourceModel $fixture } 'duplicado'

$destination = Join-Path $temp 'published.ppam'
$candidate = Join-Path $temp 'candidate.ppam'
Write-TestText $destination 'previous'
Write-TestText $candidate 'next'
$lock = [IO.File]::Open($destination,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
try { Assert-Throws { Publish-File $candidate $destination } '.*' } finally { $lock.Dispose() }
Assert-True ((Get-Content $destination -Raw) -eq 'previous') 'Locked destination preserved'
Assert-True ((Get-Content $candidate -Raw) -eq 'next') 'Candidate preserved on failure'
Publish-File $candidate $destination
Assert-True ((Get-Content $destination -Raw) -eq 'next') 'Candidate published'
$backups = @(Get-ChildItem $temp -Filter 'published.ppam.*.bak')
Assert-True ($backups.Count -eq 1) 'Backup created'
Assert-True ((Get-Content $backups[0].FullName -Raw) -eq 'previous') 'Backup contains previous version'

$package = Join-Path $temp 'fixture.ppam'
$zip = [IO.Compression.ZipFile]::Open($package,[IO.Compression.ZipArchiveMode]::Create)
try {
    Set-ZipText $zip '[Content_Types].xml' '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"/></Types>'
    Set-ZipText $zip '_rels/.rels' '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
    Set-ZipText $zip 'ppt/vbaProject.bin' 'Synthetic fixture: not an actual VBA project.'
} finally { $zip.Dispose() }
Assert-Throws { Test-PpamPackage $package } 'Parte ausente'
Add-Ribbon $package (Join-Path $repo 'resources\customUI.xml')
Test-PpamPackage $package
$script:passed++
Add-Ribbon $package (Join-Path $repo 'resources\customUI.xml')
Test-PpamPackage $package
$script:passed++
$zip = [IO.Compression.ZipFile]::Open($package,[IO.Compression.ZipArchiveMode]::Update)
try {
    Assert-True ((Read-ZipText $zip 'ppt/vbaProject.bin') -eq 'Synthetic fixture: not an actual VBA project.') 'Ribbon packaging preserves VBA bytes'
    $ct = Read-ZipText $zip '[Content_Types].xml'
    Set-ZipText $zip '[Content_Types].xml' ($ct.Replace('addin.macroEnabled','presentation.macroEnabled'))
} finally { $zip.Dispose() }
Assert-Throws { Test-PpamPackage $package } 'Formato nativo'
$zip = [IO.Compression.ZipFile]::Open($package,[IO.Compression.ZipArchiveMode]::Update)
try {
    $ct = Read-ZipText $zip '[Content_Types].xml'
    Set-ZipText $zip '[Content_Types].xml' ($ct.Replace('presentation.macroEnabled','addin.macroEnabled'))
    $zip.GetEntry('ppt/vbaProject.bin').Delete()
} finally { $zip.Dispose() }
Assert-Throws { Test-PpamPackage $package } 'Projeto VBA ausente'

$fallbackPptm = Join-Path $temp 'fallback-source.pptm'
$fallbackPpam = Join-Path $temp 'fallback-result.ppam'
$zip = [IO.Compression.ZipFile]::Open($fallbackPptm,[IO.Compression.ZipArchiveMode]::Create)
try {
    Set-ZipText $zip '[Content_Types].xml' '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"/></Types>'
    Set-ZipText $zip 'ppt/vbaProject.bin' 'VBA bytes must remain unchanged.'
} finally { $zip.Dispose() }
Convert-PptmPackageToPpam $fallbackPptm $fallbackPpam
$zip = [IO.Compression.ZipFile]::OpenRead($fallbackPpam)
try {
    $fallbackTypes = Read-ZipText $zip '[Content_Types].xml'
    Assert-True ($fallbackTypes -match 'addin\.macroEnabled\.main\+xml') 'Fallback sets PPAM content type'
    Assert-True ((Read-ZipText $zip 'ppt/vbaProject.bin') -eq 'VBA bytes must remain unchanged.') 'Fallback preserves VBA bytes'
} finally { $zip.Dispose() }
Assert-Throws { Convert-PptmPackageToPpam $fallbackPpam (Join-Path $temp 'invalid.ppam') } 'origem.*PPTM'

# Mock only process discovery; a build guard must never try to stop that process.
function Get-Process { param($Name,$ErrorAction) [pscustomobject]@{Id=123;ProcessName='POWERPNT'} }
Assert-Throws { Assert-PowerPointClosed } 'Feche o PowerPoint'
Remove-Item Function:\Get-Process
Write-Host "PASS: $script:passed checks. Fixtures retained at $temp"
