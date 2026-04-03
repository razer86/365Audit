# Module dependencies for the 365Audit Azure Function.
# All modules are installed by the managed dependency system at cold start.
# Listing Graph SDK modules here ensures assembly versions align with Az modules.
@{
    'Az.Accounts'                                = '3.*'
    'Az.KeyVault'                                = '6.*'
    'Microsoft.Graph.Authentication'             = '2.*'
    'Microsoft.Graph.Applications'               = '2.*'
    'Microsoft.Graph.DeviceManagement'           = '2.*'
    'Microsoft.Graph.DeviceManagement.Enrollment' = '2.*'
    'Microsoft.Graph.Devices.CorporateManagement' = '2.*'
    'Microsoft.Graph.Groups'                     = '2.*'
    'Microsoft.Graph.Identity.DirectoryManagement' = '2.*'
    'Microsoft.Graph.Identity.SignIns'            = '2.*'
    'Microsoft.Graph.Reports'                    = '2.*'
    'Microsoft.Graph.Users'                      = '2.*'
}
