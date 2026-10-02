param([string]$TaskName = '1CReportsDaily')
$ErrorActionPreference = 'Stop'
$ExportRoot = $PSScriptRoot
$ExportConfig = Get-Content -LiteralPath (Join-Path $ExportRoot 'bases.json') -Encoding UTF8 | ConvertFrom-Json
if (-not $ExportConfig.production_enabled) {
    throw 'The export pipeline has not passed production validation. No task was created.'
}
if ((Get-TimeZone).Id -ne 'Russian Standard Time') {
    throw 'This schedule expects Moscow time. The system time zone was not changed.'
}
$ExportPython = Join-Path $ExportRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $ExportPython)) {
    $ExportPython = [IO.Path]::GetFullPath((Join-Path $ExportRoot '..\work\.venv\Scripts\python.exe'))
}
$ExportScript = Join-Path $ExportRoot 'nightly.py'
if (-not (Test-Path -LiteralPath $ExportPython) -or -not (Test-Path -LiteralPath $ExportScript)) {
    throw 'The configured Python runtime or export entry point is missing.'
}
$ExportExisting = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -ne $ExportExisting) {
    throw 'A task with this name already exists. Inspect it before changing the schedule.'
}
$ExportIdentity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$ExportPrincipal = New-ScheduledTaskPrincipal -UserId $ExportIdentity -LogonType Interactive -RunLevel Limited
$ExportAction = New-ScheduledTaskAction -Execute $ExportPython -Argument ('-X utf8 "' + $ExportScript + '"') -WorkingDirectory $ExportRoot
$ExportFirstRun = (Get-Date).Date.AddDays(1).AddHours(3)
$ExportTrigger = New-ScheduledTaskTrigger -Daily -At $ExportFirstRun
$ExportSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 6) -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 10)
Register-ScheduledTask -TaskName $TaskName -Action $ExportAction -Trigger $ExportTrigger -Settings $ExportSettings -Principal $ExportPrincipal -Description 'Export the previous Moscow calendar day from 1C to the configured Google Drive folder.'
# Interactive logon avoids storing the Windows password. The Windows account
# must remain signed in; Codex does not need to be running.
