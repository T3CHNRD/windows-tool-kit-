@{
    RootModule = 'MaintenanceToolkit.psm1'
    ModuleVersion = '1.0.0'
    GUID = '0e21f665-9c4b-471f-aa12-d31ef0433eb9'
    Author = 'Toolkit Builder'
    CompanyName = 'Personal Toolkit'
    Copyright = '(c) 2026'
    Description = "Core module for T3CHNRD'S Windows Tool Kit."
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Get-ToolkitRoot',
        'Get-ToolkitSettings',
        'Test-ToolkitIsAdmin',
        'Write-ToolkitLog',
        'New-ToolkitTaskContext',
        'Get-ToolkitTaskCatalog',
        'Invoke-ToolkitTaskById'
    )
    CmdletsToExport = @()
    VariablesToExport = '*'
    AliasesToExport = @()
}
