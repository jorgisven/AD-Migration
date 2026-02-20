@{
    RootModule        = 'ADMigration.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'bf9e8f4d-9a4e-4c2f-9c1a-123456789abc'
    Author            = 'jorgisven'
    CompanyName       = 'Internal'
    Copyright         = '(c) 2026'
    Description       = 'Module for SAICPRINT → CRIT.AD migration tasks, including exports, transforms, imports, and logging.'

    FunctionsToExport = @('Write-Log')
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags = @('ActiveDirectory','Migration','GPO','OU','Logging')
        }
    }
}