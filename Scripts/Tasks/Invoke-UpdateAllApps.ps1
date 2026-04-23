[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Output 'Updating all supported apps with winget...'
winget.exe upgrade --all --include-unknown --accept-source-agreements --accept-package-agreements | Out-String | Write-Output
Write-Output 'App update run completed.'
