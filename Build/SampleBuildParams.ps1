$params = @{
    Location             = 'EastUS'
    RGName               = 'SPOPermission-RG-01'
    FunctionAppName      = 'func-SPOPerm-01-20220530'
    StorageAccountName   = 'safuncspoperm0120220530'
    KeyVaultName         = 'kvSPOPerm0120220530'
    PnPApplicationName   = "func-ecl-SPOPerm-01-PnPApp"
    LogAnalyticsMaxLevel = 5
    CreateTestSP         = $true
}
$Location = 'EastUS'
$RGName = 'SPOPermission-RG-01'
$FunctionAppName = 'func-SPOPerm-01-20220530'
$StorageAccountName = 'safuncspoperm0120220530'
$KeyVaultName = 'kvSPOPerm0120220530'
$PnPApplicationName = "func-ecl-SPOPerm-01-PnPApp"
$LogAnalyticsMaxLevel = 5
$CreateTestSP = $true
