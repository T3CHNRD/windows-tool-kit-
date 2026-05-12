Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ToolkitRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:TaskRoot = Join-Path $script:ToolkitRoot 'Scripts\Tasks'
. (Join-Path $script:TaskRoot 'VendorUpdate.Common.ps1')

function Get-TtkSupportedUpdateManufacturers {
    return @('Dell', 'HP', 'Lenovo', 'Framework')
}

function Invoke-TtkBiosUpdate {
    [CmdletBinding()]
    param()
    Invoke-VendorMaintenanceUpdate -Mode BIOS -PromptForManufacturer
}

function Invoke-TtkFirmwareUpdate {
    [CmdletBinding()]
    param()
    Invoke-VendorMaintenanceUpdate -Mode Firmware -PromptForManufacturer
}

function Invoke-TtkDriverUpdate {
    [CmdletBinding()]
    param()
    Invoke-VendorMaintenanceUpdate -Mode Drivers -PromptForManufacturer
}

function Invoke-TtkWindowsUpdate {
    [CmdletBinding()]
    param(
        [string]$SkipSelectionFile
    )
    Invoke-WindowsUpdateInstallation -SkipSelectionFile $SkipSelectionFile
}

function Invoke-TtkAppUpdate {
    [CmdletBinding()]
    param()
    & (Join-Path $script:TaskRoot 'Invoke-UpdateAllApps.ps1')
}

Export-ModuleMember -Function @(
    'Get-TtkSupportedUpdateManufacturers',
    'Invoke-TtkBiosUpdate',
    'Invoke-TtkFirmwareUpdate',
    'Invoke-TtkDriverUpdate',
    'Invoke-TtkWindowsUpdate',
    'Invoke-TtkAppUpdate'
)
