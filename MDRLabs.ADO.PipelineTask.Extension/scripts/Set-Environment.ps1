$npm_cfg_ado_pat = 'your-pat-goes-here'
$npm_cfg_ado_org = 'your-organization-goes-here'
$npm_cfg_ado_ext_publisher = 'your-publisher-id-goes-here'
$npm_cfg_ado_ext_developer = 'your-name-goes-here'
$collectionuri = 'https://dev.azure.com/your-organization-goes-here/'

[Environment]::SetEnvironmentVariable("npm_cfg_ado_pat", $npm_cfg_ado_pat, "User")
[Environment]::SetEnvironmentVariable("npm_cfg_ado_org", $npm_cfg_ado_org, "User")
[Environment]::SetEnvironmentVariable("npm_cfg_ado_ext_publisher", $npm_cfg_ado_ext_publisher, "User")
[Environment]::SetEnvironmentVariable("npm_cfg_ado_ext_developer", $npm_cfg_ado_ext_developer, "User")
[Environment]::SetEnvironmentVariable("SYSTEM_COLLECTIONURI", $collectionuri, "User")