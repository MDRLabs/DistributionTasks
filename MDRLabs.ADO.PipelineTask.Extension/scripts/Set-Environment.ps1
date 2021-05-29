$pat = 'qm72xixg5jwwa4klij5swrb6g3vejj2atqys7uzi3es4jyqn7q2a'
$VSTSAccount = 'cbsp-abnamro'
$Publisher = 'ABNAMRO-DES-AZURE'
$VSTSExtensionDeveloper = 'Rohit Nadhe'
$collectionuri = 'https://dev.azure.com/cbsp-abnamro/'

[Environment]::SetEnvironmentVariable("npm_config_vsts_pat", $pat, "User")
[Environment]::SetEnvironmentVariable("VSTSAccount", $VSTSAccount, "User")
[Environment]::SetEnvironmentVariable("Publisher", $Publisher, "User")
[Environment]::SetEnvironmentVariable("VSTSExtensionDeveloper", $VSTSExtensionDeveloper, "User")
[Environment]::SetEnvironmentVariable("SYSTEM_COLLECTIONURI", $collectionuri, "User")