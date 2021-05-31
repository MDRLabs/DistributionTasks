$pat = 'your-pat-goes-here'
$ADOAccount = 'your-organization-goes-here'
$Publisher = 'your-publisher-id-goes-here'
$ADOExtensionDeveloper = 'your-name-goes-here'
$collectionuri = 'https://dev.azure.com/your-organization-goes-here/'

[Environment]::SetEnvironmentVariable("npm_config_ado_pat", $pat, "User")
[Environment]::SetEnvironmentVariable("ADOAccount", $ADOAccount, "User")
[Environment]::SetEnvironmentVariable("Publisher", $Publisher, "User")
[Environment]::SetEnvironmentVariable("ADOExtensionDeveloper", $ADOExtensionDeveloper, "User")
[Environment]::SetEnvironmentVariable("SYSTEM_COLLECTIONURI", $collectionuri, "User")