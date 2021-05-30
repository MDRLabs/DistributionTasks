<#
    .Synopsis
        Helper functions related to Azure DevOps REST APIs
    
    .Description
        Functions for manipulating JSON files and invoking REST APIs

    .Notes
        The Module is dependent upon certain environment variables. Namely:
          1. $env:SYSTEM_COLLECTIONURI is the GUID of the Azure DevOps organisation
          2. $env:SYSTEM_ACCESSTOKEN is an agent-scoped variable containing an OAuth token, generated when
            the build pipeline option "Allow script access to OAuth token" is enabled
        
        Due to the change in the URL used by Azure DevOps it is necessary to cater for both the old
        VSTS url as well as the new Azure DevOps one. The Vssps endpoint uses the new url
#>

Set-StrictMode -Version Latest

[string]$ApiVersion4            = "api-version=5.1-preview.1"

[string]$ApiTaskGroups          = "api-version=5.1-preview.1"
[string]$ApiVariableGroups      = "api-version=5.1-preview.1"
[string]$ApiGraphGroups         = "api-version=5.1-preview.1"
[string]$ApiAccessController    = "api-version=5.1"
[string]$ApiSecurityNamespaces  = "api-version=5.1"

[string]$ApiServiceEndpoints    = "api-version=5.1-preview.2"
[string]$ApiAzureDevOps         = "api-version=5.1"

[string]$OrganisationUrl = $env:SYSTEM_COLLECTIONURI 
[string]$Organisation = $OrganisationUrl -replace '^https:\/\/(?:(?:dev.azure.com\/)(.*?)\/|(.*?)(?:\.visualstudio\.com\/))', '$1$2'
[string]$Vssps = "https://vssps.dev.azure.com/${Organisation}/"
[string]$Base64AuthInfo = $null
[string]$AuthorisationType = "Basic"
[string]$SonarQubePrepareAnalysisConfigTaskId = "15b84ca1-b62f-4a2a-a403-89b77a063157"
[string]$SplunkRestCallTaskId = "d8a3d2d0-20f9-11e7-b752-57165c2d4193"
[int]$MetaTaskSecurityNamespaceActionBit = 0
[System.Collections.Hashtable]$MetaTaskSecurityNamespace = New-Object 'System.Collections.Hashtable'
[psobject]$Groups = $null
[System.Collections.Generic.Dictionary[string, string]]$ProjectNameIdLookup = New-Object 'System.Collections.Generic.Dictionary[string,string]'
[System.Collections.Generic.Dictionary[string, string]]$ProjectProperties = New-Object 'System.Collections.Generic.Dictionary[string,string]'

function Get-OrganizationName
{
    <#
        .Synopsis
            Return the current organization name.
    #>

    $script:Organisation
}

function Get-TaskCollections{
    <#
        .Synopsis
            Return the key value pair of service connection task collection.
    #>

    $TaskCollection = @{
        'SonarQube' = '15b84ca1-b62f-4a2a-a403-89b77a063157';
        'Splunk' = 'd8a3d2d0-20f9-11e7-b752-57165c2d4193'
    }

    $TaskCollection
}

function Get-SonarQubePrepareAnalysisConfigTaskId {
    <#
        .Synopsis
            Return the task id of the SonarQube Prepare Analysis Configuration
            task.
    #>
    $script:SonarQubePrepareAnalysisConfigTaskId
}

function Get-SplunkRestCallTaskId {
    <#
        .Synopsis
            Return the task id of the Splunk Prepare Analysis Configuration
            task.
    #>
    $script:SplunkRestCallTaskId
}

function Get-Base64AuthInfo {
    <#
        .Synopsis
            Return a Basic Auth Credential
    #>
    $script:Base64AuthInfo
}

function Get-AuthorisationType {
    <#
        .Synopsis
            Return the current Authentication Method
    #>
    $script:AuthorisationType
}

function Get-TaskgroupPostUrl {
    <#
        .Synopsis
            Return the taskgroup POST url for the specified team project within the current organisation
            POST https://dev.azure.com/{organization}/{project}/_apis/distributedtask/taskgroups?api-version=5.1-preview.1

        .Parameter Project
            The team project name
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Task group add REST endpoint expects the team project name. It is null or empty"
    }
    "${script:OrganisationUrl}$Project/_apis/distributedtask/taskgroups?$script:ApiTaskGroups"
}

function Get-TaskgroupGetUrl {
    <#
        .Synopsis
            If a Taskgroup Id has been provided, return the Taskgroup GET URL for the specific Taskgroup. Otherwise return the URL to
            retrieve all Task groups for the specified team project in the current organisation.
            GET https://dev.azure.com/{organization}/{project}/_apis/distributedtask/taskgroups/{taskGroupId}?api-version=5.1-preview.1 

        .Parameter Project
            The team project name

        .Parameter TaskgroupId
        .Parameter Id
            The uuid of a specific Taskgroup instance in the specified team project

        .Parameter DisablePriorVersions
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project,
        [Parameter(Mandatory = $false)][string]$TaskgroupId = "",
        [Parameter(Mandatory = $false)][string]$DisablePriorVersions = ""
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Task group Get REST endpoint expects the team project name. It is null or empty"
    }

    [string]$restUrl = $null

    if (-not ([string]::IsNullOrEmpty($TaskgroupId))) {
        $restUrl = "${script:OrganisationUrl}$Project/_apis/distributedtask/taskgroups/${TaskgroupId}?$script:ApiTaskGroups"
        if (-not ([string]::IsNullOrEmpty($DisablePriorVersions))) {
            $restUrl = "${script:OrganisationUrl}$Project/_apis/distributedtask/taskgroups/${TaskgroupId}?disablePriorVersions=${DisablePriorVersions}&$script:ApiTaskGroups"
        }
    }
    else {
        $restUrl = "${script:OrganisationUrl}$Project/_apis/distributedtask/taskgroups?$script:ApiTaskGroups"
    }
    $restUrl
}

function Get-TaskgroupPutUrl {
    <#
        .Synopsis
            Return the Taskgroup Update URL for this organisation and the specified Taskgroup Id
            PUT https://dev.azure.com/{organization}/{project}/_apis/distributedtask/taskgroups/{taskGroupId}?api-version=5.1-preview.1

        .Parameter Project
            The team project name

        .Parameter ParentTaskgroupId
            In case the taskgroup has to be published as preview
        .Parameter TaskgroupId
            In case the taskgroup has to be updated ultimately incrementing only the minor version
        .Parameter Id
            The uuid of a specific Taskgroup instance in the specified team project
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project,
        [Parameter(Mandatory = $false)][string]$ParentTaskgroupId,
        [Parameter(Mandatory = $false)][string]$TaskgroupId
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Task group update REST endpoint expects the team project name. It is null or empty"
    }

    [string]$restUrl = $null

    if (-not ([string]::IsNullOrEmpty($ParentTaskgroupId))) {
        $restUrl = "${script:OrganisationUrl}$Project/_apis/distributedtask/taskgroups?parentTaskGroupId=${ParentTaskgroupId}&$script:ApiTaskGroups"
    }
    elseif(-not ([string]::IsNullOrEmpty($TaskgroupId))) {    
        $restUrl = "${script:OrganisationUrl}$Project/_apis/distributedtask/taskgroups?TaskGroupId=${TaskgroupId}&$script:ApiTaskGroups"
    }
    else {
        throw "The Task group update REST endpoint expects a parent task group id. ParentTaskgroupId is null or empty"
    }
    $restUrl
}

function Get-VariablegroupGetUrl {
    <#
        .Synopsis
            Return the Variablegroup GET URL for the specific variable group name.
            GET https://dev.azure.com/{organization}/{project}/_apis/distributedtask/variablegroups?groupName={groupnames}&api-version=5.1-preview.1 

        .Parameter Project
            The team project name

        .Parameter GroupName
             The variable group name
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Project,

        [Parameter(Mandatory = $true)]
        [string]$GroupName
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Variable group Get REST endpoint expects the team project name. It is null or empty"
    }

    if ([string]::IsNullOrEmpty($GroupName)) {
        throw "The Variable group Get REST endpoint expects the group name. It is null or empty"
    }

    "${script:OrganisationUrl}$Project/_apis/distributedtask/variablegroups?groupName=$GroupName&$script:ApiVariableGroups"
}

function Get-VariablegroupPostUrl {
    <#
        .Synopsis
            Return the Variablegroup POST URL for the specific project name. 
            POST https://dev.azure.com/{organization}/{project}/_apis/distributedtask/variablegroups?api-version=5.1-preview.1 

        .Parameter Project
            The team project name
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Project
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Variable group Get REST endpoint expects the team project name. It is null or empty"
    }

    "${script:OrganisationUrl}$Project/_apis/distributedtask/variablegroups?$script:ApiVariableGroups"
}

function Get-VariablegroupPutUrl {
    <#
        .Synopsis
            Return the Variablegroup PUT URL for project and variable group name.
            PUT https://dev.azure.com/{organization}/{project}/_apis/distributedtask/variablegroups/{VariableGroupObjectId}?api-version=5.1-preview.1 

        .Parameter Project
            The team project name

        .Parameter VariableGroupObjectId
            The variable group object id
    #>
    param (
        [Parameter(Mandatory = $true)]
        [string]$Project,

        [Parameter(Mandatory = $true)]
        [int]$VariableGroupObjectId
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Variable group Get REST endpoint expects the team project name. It is null or empty"
    }

    if ($VariableGroupObjectId -le 0) {
        throw "The Variable group Get REST endpoint expects the variable group object id. It is null or 0"
    }


    "${script:OrganisationUrl}$Project/_apis/distributedtask/variablegroups/${VariableGroupObjectId}?$script:ApiVariableGroups"
}

function Get-ServiceConnectionGetUrl {
    <#
        .Synopsis
            Return the Service connection GET URL for the specific Service connection endpoint otherwise return all service connections for current organizations team project.
            GET https://dev.azure.com/{organization}/{project}/_apis/serviceendpoint/endpoints?endpointId={endpointId}&api-version=5.1-preview.2 
            GET https://dev.azure.com/{organization}/{project}/_apis/serviceendpoint/endpoints?api-version=5.1-preview.2
        .Parameter Project
            The team project name

        .Parameter EndpointId       
            The service connection end point id
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project,
        [Parameter(Mandatory = $false)][string]$EndpointId = ''
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Service connection Get REST endpoint expects the team project name. It is null or empty"
    }

    [string]$restUrl = $null

    if (-not [string]::IsNullOrEmpty($EndpointId)) {
        $restUrl ="${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints?endpointId=${EndpointId}&$script:ApiServiceEndpoints" 
    }
    else {
        $restUrl ="${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints?$script:ApiServiceEndpoints"    
    }
    $restUrl
}

function Get-ServiceConnectionGetByNameUrl {
    <#
        .Synopsis
            Return the Service connection GET URL for the specific Service connection endpoint otherwise return all service connections for current organizations team project.
           GET https://dev.azure.com/{organization}/{project}/_apis/serviceendpoint/endpoints?endpointNames={endpointNames}&api-version=5.1-preview.2
        .Parameter Project
            The team project name

        .Parameter EndpointName    
            The service connection end point names list
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project,
        [Parameter(Mandatory = $true)][string]$EndpointName
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Service connection Get REST endpoint expects the team project name. It is null or empty"
    }

    if ($EndpointName.Length -eq 0) {
        throw "The Service connection Get REST endpoint expects endpoint names list. It is null or empty"
    }

    "${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints?endpointNames=$EndpointName&$script:ApiServiceEndpoints" 
}

function Get-ServiceConnectionPostUrl {
    <#
        .Synopsis
            Return the service connection POST url for the specified team project within the current organisations
            POST https://dev.azure.com/{organization}/{project}/_apis/serviceendpoint/endpoints?api-version=5.1-preview.2

        .Parameter Project
            The team project name
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Service connection add REST endpoint expects the team project name. It is null or empty"
    }
    "${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints?$script:ApiServiceEndpoints"
}

function Get-ServiceConnectionsPutUrl {
    <#
        .Synopsis
            Return the service connection Update URL for this organisation and project
            PUT https://dev.azure.com/{organization}/{project}/_apis/serviceendpoint/endpoints?api-version=5.1-preview.2
        .Parameter Project
            The team project name
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The service connection update REST endpoint expects the team project name. It is null or empty"
    }

    "${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints?$script:ApiServiceEndpoints"
}

function Get-GraphGroupsUrl {
    <#
        .Synopsis
            Return a Graph URL for groups
            GET https://vssps.dev.azure.com/{organization}/_apis/graph/groups/{groupDescriptor}?api-version=5.1-preview.1
        
        .Parameter Token
            Continuation token. Since the list of groups may be large, results are returned in pages of groups. 
            If there are more results than can be returned in a single page, the result set will containt a continuation token
            for retrieval of the next set of results.
    #>
    param (
        [Parameter(Mandatory = $false)][string]$Token = ""
    )
    
    [string]$restUrl = $null

    if ([string]::IsNullOrEmpty($Token)) {
        $restUrl = "${script:Vssps}_apis/graph/groups?subjectTypes=vssgp&$script:ApiGraphGroups"
    }
    else {
        $restUrl = "${script:Vssps}_apis/graph/groups?subjectTypes=vssgp&continuationToken=$($Token)&$script:ApiGraphGroups"
    }
    $restUrl
}

function Get-SecurityNamespacesUrl {
    <#
        .Synopsis
            Return a Security-Namespaces URL
            GET https://dev.azure.com/{organization}/_apis/securitynamespaces/{securityNamespaceId}?localOnly={localOnly}&api-version=5.1
    #>
    "${script:OrganisationUrl}_apis/securitynamespaces?$script:ApiSecurityNamespaces"
}

function Set-MetaTaskSecurityNameSpace {
    <#
        Create a PSObject containing the Azure DevOps Security Namespace for MetaTask. This repesents the 
        security namespace for Taskgroups
    #>
    [string]$resturl = Get-SecurityNamespacesUrl
    [psobject]$securityNamespaces = Invoke-RestMethod -Uri $resturl -Method Get -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
    [psobject]$o = ($securityNamespaces.value | Where-Object -FilterScript {$_.name -eq 'MetaTask'})
    foreach ($property in $o.psobject.properties.name) {
        $script:MetaTaskSecurityNamespace[$property] = $o.$property
    }
    $script:MetaTaskSecurityNamespace
}

function Get-AclUrl {
    <#
        Return the Access Control List (ACL) for a specified Security Namespace, Project and Taskgroup
        GET https://dev.azure.com/{organization}/_apis/accesscontrollists/{securityNamespaceId}?api-version=5.1
        GET https://dev.azure.com/{organization}/_apis/accesscontrollists/{securityNamespaceId}?token={token}&descriptors={descriptors}&includeExtendedInfo={includeExtendedInfo}&recurse={recurse}&api-version=5.1
    #>
    param (
        [Parameter(Mandatory = $true)][string]$SecurityNamespaceId,
        [Parameter(Mandatory = $false)][string]$ProjectId = "",
        [Parameter(Mandatory = $false)][string]$TaskgroupId = ""
    )

    if ([string]::IsNullOrEmpty($SecurityNamespaceId)) {
        throw "The Access Control Lists REST endpoint expects the Security Namespace Id. It is null or empty"
    }

    if (-not [string]::IsNullOrEmpty($ProjectId) -and -not [string]::IsNullOrEmpty($TaskgroupId)) {
        "${script:OrganisationUrl}_apis/accesscontrollists/$($SecurityNamespaceId)?token=$($ProjectId)/$($TaskgroupId)&$script:ApiAccessController"
    }
    elseif (-not [string]::IsNullOrEmpty($ProjectId) -and [string]::IsNullOrEmpty($TaskgroupId)) {
        "${script:OrganisationUrl}_apis/accesscontrollists/$($SecurityNamespaceId)?token=$($ProjectId)&$script:ApiAccessController"
    }
    else {
        "${script:OrganisationUrl}_apis/accesscontrollists/$($SecurityNamespaceId)?$script:ApiAccessController"
    }
}

function Get-Identity {
    <#
        .Synopsis
            Return Get the descriptor of the Security Group
            GET https://vssps.dev.azure.com/{organization}/_apis/graph/descriptors/{storageKey}?api-version=5.1-preview.1
    #>
    param (
        [Parameter(Mandatory = $true)][string]$GroupId
    )
    
    if ([string]::IsNullOrEmpty($GroupId)) {
        throw "The Identities REST endpoint expects the Groud Id. It is null or empty"
    }

    "${script:Vssps}_apis/Identities/$($GroupId)?$script:ApiVersion4"
}

function Get-ValidateJsonSchema
{
    <#
        .Synopsis
            Evaluate if the file passed to the script is valid Json 
        
        .Parameter JsonText
            The relative json text
    #>
    param (
        [Parameter(Mandatory = $true)][string]$JsonText
    )

    try {
        $JsonObject = ConvertFrom-Json -InputObject $JsonText -ErrorAction Stop
        $o = [PSCustomObject]@{
            JsonObject  = $JsonObject
        }
    }
    catch {
        $o = [PSCustomObject]@{
            JsonObject  = $null
        }
    }
    $o
}

function Get-JsonContent {
    <#
        .Synopsis
            Evaluate if the file passed to the script is valid Json and that the expected
            API type is present.
        
        .Parameter Filename
        .Parameter File
            The relative or fullname of a JSON file

        .Parameter DefinitionType
        .Parameter Type
            The definition type of the JSON object. For a task group it should be 'metaTask'

        .Parameter PropertyName
        .Parameter Name
            The name of a property that is expected. For a variable group object the property to expect is 'variables'

        .Outputs
            Return a PSObject with three members:
              JsonObject: contains the JSON 
              Name: contains the value of the JsonObject name property
              Valid: true or false whether the JSON is valid
    #>
    param (
        [Parameter(Mandatory = $true)][string][Alias("File")]$Filename,
        [Parameter(Mandatory = $true, ParameterSetName = 'DefinitionType')][Alias('Type')][string]$DefinitionType,
        [Parameter(Mandatory = $true, ParameterSetName = 'Property')][Alias('Name')][string]$PropertyName
    )
    try {
        $JsonRaw = Get-Content $Filename -Raw
        $JsonObject = ConvertFrom-Json -InputObject $JsonRaw -ErrorAction Stop
        $o = [PSCustomObject]@{
            ValidJson   = $false
            JsonContent = $JsonRaw
            JsonObject  = $JsonObject
            Name        = $JsonObject.name
        }

        if ($PSCmdlet.ParameterSetName -eq 'DefinitionType' -and $JsonObject.definitionType -eq $DefinitionType) {
            $o.ValidJson = $true
        }
        elseif ($PSCmdlet.ParameterSetName -eq 'Property') {
            foreach ($element in $JsonObject.PSobject.Properties) 
            { 
                if($element.Name -eq $PropertyName)
                {
                    $o.ValidJson = $true
                    break
                }
            }
        }
    }
    catch {
        $o = [PSCustomObject]@{
            ValidJson   = $false
            JsonContent = $null
            JsonObject  = $null
            Name        = $null
        }
    }
    $o
}

function Set-SystemAccessToken {
    <#
        .Synopsis
            To enable the script to use the build pipeline OAuth token, go to the Options tab 
            of the build pipeline and select Allow Scripts to Access OAuth Token.
            The script can use to SYSTEM_ACCESSTOKEN environment variable to access the 
            Azure Pipelines REST APIs, instead of using Basic Authentication.
        
        .Notes
            When using a Bearer token the identity that invokes rest calls must have the
            correct permissions for the action. By default the Project Collection Build Service
            identity runs a build task. It doesn't have the correct permissions. Favour the use of a 
            Personal Access Token saved as a secret process variable

        .Notes
            Sets:
              AuthorisationType: Bearer
              Base64AuthInfo: A Bearer Authorisation Token (System.AccessToken)
    #>
    $script:Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}" -f $env:SYSTEM_ACCESSTOKEN)))
    $script:AuthorisationType = "Bearer"
}

function Set-PersonalAccessToken {
    <#
        .Synopsis
            Saves the supplied Personal Access Token as a Base64String 
        
        .Parameter Token
            The user's Personal Access Token (PAT) with the correct scopes for creating and updating Taskgroups.

        .Notes
            The username is already part of the Personal Access Token. The two part Basic Authorization username:password 
            can be reduced to :token becuase in this instance username must be an empty string
            Sets:
              AuthorisationType: Basic
              Base64AuthInfo: A Basic Authorisation Token (PAT)
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Token
    )

    if ([string]::IsNullOrEmpty($Token)) {
        throw "Personal Access Token required. The supplied token is invalid"
    }

    $script:Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes((":{1}" -f "", $Token)))
}

function Get-WinCredential {
    <#
        .Synopsis
            Create a PSCredential object and use it to create a Base64Auth token.
            If a username and token are provided use them and do not store them as a Windows Credential.
            If no username and token are provided, retrieve the Generic Windows Credential for the Azure DevOps Account, if it exists,
            otherwise prompt the user for Credentials.
        
        .Parameter Username
        .Parameter User
            The user that has permission to create Taskgroups in each Team project

        .Parameter Token
            The user's Personal Access Token (PAT) with the correct scopes for creating and updating Taskgroups

        .Outputs
            Return a PSObject with three members:
              Credential: A PSCredential object 
              Base64AuthInfo: A Basic Auth token
              StoreCredential: boolean, $false
    #>
    param (
        [Parameter(Mandatory = $false)][Alias("User")][string]$Username = "",
        [Parameter(Mandatory = $false)][securestring]$Token = ""
    )

    [bool]$storeCredential = $false
    [System.Management.Automation.PSCredential]$winCredential = $null

    if ([string]::IsNullOrEmpty($Username) -or [string]::IsNullOrEmpty($Token)) {
        $winCredential = Get-StoredCredential -Type Generic -Target $script:OrganisationUri
        if ($null -eq $winCredential) {
            $winCredential = Get-Credential -Message "Azure DevOps Username and PAT"
            $storeCredential = $true
        }
    }
    else {
        $winCredential = New-Object System.Management.Automation.PSCredential($Username, $Token)
    }

    $script:Base64AuthInfo = [Convert]::ToBase64String(
        [Text.Encoding]::ASCII.GetBytes(
            (" {0}:{1}" -f $winCredential.UserName, $winCredential.GetNetworkCredential().Password )
        )
    )
    $o = [PSCustomObject]@{
        Credential      = $winCredential
        Base64AuthInfo  = $script:Base64AuthInfo
        StoreCredential = $storeCredential
    }
    $o
}

function Get-NewWinCredential {
    <#
        .Synopsis
            Save the user credential as a Windows GENERIC credential only if it is valid. The caller
            must validate the user credential
        
        .Parameter Credential
        .Parameter Cred
            The current PSCredential object that

        .Parameter Validity
            True if the Credential was successfully used to connect to an endpoint, otherwise false
    #>
    param (
        [Parameter(Mandatory = $true)][Alias("Cred")][PSCustomObject]$Credential,
        [Parameter(Mandatory = $true)][bool]$Validity
    )

    if ([string]::IsNullOrEmpty($Credential)) {
        throw "No credential was provided to store. Call the function with a valid PSCustomObject containing a Credential"
    }

    if ($Credential.StoreCredential -and $Validity ) {
        $null = New-StoredCredential -Target $script:OrganisationUri -Credentials $Credential.Credential -Type Generic -Persist LocalMachine -Comment "Azure DevOps Credential for invoking REST APIs"
    }
}

function Get-ServiceConnectionsList {
    <#
        .Synopsis
            Return a collection of service connections for the specified type. If no type is specified, 
            return all service connections for the current team project

        .Parameter Project
            The team project name

        .Parameter Type
            The service connection type. For example SonarQube

        .Outputs
            Returns a string containing a valid Azure DevOps REST API endpoint
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project,
        [Parameter(Mandatory = $false)][string]$Type = ""
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The Service Connections Get REST endpoint expects the team project name. It is null or empty"
    }

    [string]$resturl = $null

    if (-not ([string]::IsNullOrEmpty($Type))) {
        if($Type -eq 'Splunk')
        {
            $Type = 'Generic'
        }
        $resturl = "${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints?type=$Type&$script:ApiServiceEndpoints"
    }
    else {
        $resturl = "${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints?$script:ApiServiceEndpoints"
    }
    $resturl
}

function Get-ServiceconnectionTaskId{
    <#
       .Synopsis
           Get the Service connection task id.

       .Parameter ServiceConnectionType
       .Parameter SC
           A service connection type

       .Outputs
           Returns a string containing a valid task id for service connection
   #>
   param (
       [Parameter(Mandatory = $true)][Alias("SC")][string]$ServiceConnectionType
   )
   
   [string]$taskId = $null
   switch ( $ServiceConnectionType ) {
       'SonarQube'   { 
           $taskId = $script:SonarQubePrepareAnalysisConfigTaskId
       }
       'Splunk'      { 
           $taskId = $script:SplunkRestCallTaskId
       }
   }

   $taskId
}

function Get-ServiceConnectionUuid {
    <#
        .Synopsis
            Get the Service connection UUID for the production instance, given a valid JSON object
            containing a collection of instances.

        .Parameter ServiceConnectionType
        .Parameter SCT
            A service Connections type for a specific team project
        
        .Parameter ServiceConnections
        .Parameter SC
            A PSObject containing one or more Service Connections for a specific team project

        .Outputs
            Returns a string containing a valid Azure DevOps REST API endpoint
    #>
    param (
        [Parameter(Mandatory = $true)][Alias("SCT")][string]$ServiceConnectionType,
        [Parameter(Mandatory = $true)][Alias("SC")][psobject]$ServiceConnections
    )

    $serviceConnectionTypeToLower = $ServiceConnectionType.ToLower()
    [string]$endpointPattern = "(?i)$serviceConnectionTypeToLower(?:\s+|[-_])(?:pr)(?:oduction){0,1}"
    [string]$uuid = $null

    foreach ($connection in $ServiceConnections.value) {
        if ($connection.name -match $endpointPattern) {
            $uuid = $connection.id
        }
    }
    if ($null -eq $uuid) {
        throw "No $ServiceConnectionType Production service connection was found. If one is expected, create it first and then re-run."
    }
    
    $uuid
}

function Get-ServiceConnectionPutUrl {
    <#
        .Synopsis
            Return the service connection Update URL for this organisation and the specified team project and endpoint
            PUT https://dev.azure.com/{organization}/{project}/_apis/serviceendpoint/endpoints/{endpointId}?api-version=5.1-preview.2
            PUT https://dev.azure.com/{organization}/{project}/_apis/serviceendpoint/endpoints/{endpointId}?operation={operation}&api-version=5.1-preview.2
        .Parameter Project
            The team project name

        .Parameter EndpointId
            The endpoint id

        .Parameter Operation    
            The operation
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project,
        [Parameter(Mandatory = $true)][string]$EndpointId,
        [Parameter(Mandatory = $false)][string]$Operation = ""
    )

    if ([string]::IsNullOrEmpty($Project)) {
        throw "The service connection update REST endpoint expects the team project name. It is null or empty"
    }

    if ([string]::IsNullOrEmpty($EndpointId)) {
        throw "The service connection update REST endpoint expects the endpoint id. It is null or empty"
    }

    [string]$restUrl = $null

    if ([string]::IsNullOrEmpty($Operation)) {
        $restUrl = "${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints/${EndpointId}?$script:ApiServiceEndpoints"
    }
    else {
        $restUrl = "${script:OrganisationUrl}$Project/_apis/serviceendpoint/endpoints/${EndpointId}?operation=${Operation}&$script:ApiServiceEndpoints"
    }
    
    $restUrl
}

function Get-TaskInputKey{
    <#
       .Synopsis
           Get the service connection task input key.

       .Parameter ServiceConnectionType
       .Parameter SC
           A service connection type

       .Outputs
           Returns a string containing a valid task id for service connection
   #>
   param (
       [Parameter(Mandatory = $true)][Alias("SC")][string]$ServiceConnectionType
   )
   
   [string]$taskInputKey = $null
   switch ( $ServiceConnectionType ) {
       'SonarQube'   { $taskInputKey = 'SonarQube' }
       'Splunk'      { $taskInputKey = 'webserviceEndpoint' }
   }

   $taskInputKey
}

function Get-SonarQubeServiceConnectionUuid {
    <#
        .Synopsis
            Get the SonarQube Service connection UUID for the production instance, given a valid JSON object
            containing a collection of instances.

        .Parameter ServiceConnections
        .Parameter SC
            A PSObject containing one or more Service Connections for a specific team project

        .Outputs
            Returns a string containing a valid Azure DevOps REST API endpoint
    #>
    param (
        [Parameter(Mandatory = $true)][Alias("SC")][psobject]$ServiceConnections
    )

    [string]$sonarEndpointPattern = "(?i)sonarqube(?:\s+|[-_])(?:pr)(?:oduction){0,1}"
    [string]$uuid = $null

    foreach ($connection in $ServiceConnections.value) {
        if ($connection.name -match $sonarEndpointPattern) {
            $uuid = $connection.id
        }
    }
    if ($null -eq $uuid) {
        throw "No SonarQube Production service connection was found. If one is expected, create it first and then re-run."
    }
    $uuid
}

function Get-SplunkServiceConnectionUuid {
    <#
        .Synopsis
            Get the Splunk Service connection UUID for the production instance, given a valid JSON object
            containing a collection of instances.

        .Parameter ServiceConnections
        .Parameter SC
            A PSObject containing one or more Service Connections for a specific team project

        .Outputs
            Returns a string containing a valid Azure DevOps REST API endpoint
    #>
    param (
        [Parameter(Mandatory = $true)][Alias("SC")][psobject]$ServiceConnections
    )

    [string]$splunkEndpointPattern = "(?i)splunk(?:\s+|[-_])(?:pr)(?:oduction){0,1}"
    [string]$uuid = $null

    foreach ($connection in $ServiceConnections.value) {
        if ($connection.name -match $splunkEndpointPattern) {
            $uuid = $connection.id
        }
    }
    if ($null -eq $uuid) {
        throw "No Splunk Production service connection was found. If one is expected, create it first and then re-run."
    }
    $uuid
}

function Get-AllAzureDevOpsTeamProjectNames {
    <#
        .Synopsis
            Initialise and return a Dictionary of Azure DevOps Team Project names and ids
            GET https://dev.azure.com/{organization}/_apis/projects?api-version=5.1

        .Outputs
            Returns a Dictionary containing the name and id of all the Azure DevOps teams for this organisation
    #>
    if ([string]::IsNullOrEmpty($script:Base64AuthInfo)) {
        throw "An Access Token is required for Authentication"
    }
    
    [string]$getAllTeamProjectsUrl = "${script:OrganisationUrl}_apis/projects?$script:ApiAzureDevOps&`$top=900"

    [psobject]$response = Invoke-RestMethod -Uri $getAllTeamProjectsUrl -Method GET -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
    foreach ($project in $response.value) {
        $ProjectNameIdLookup.Add($project.name, $project.id)
    }
    $ProjectNameIdLookup
}

function Get-Groups {
    <#
        .Synopsis
            Creates a collection of security groups for the current Azure DevOps organisation
        
        .Notes
            With PowerShell 6.x it's possible to simplify the code because of extensions to the Invoke-RestMethod that supports pagination
            and processing Response Headers. See https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-restmethod?view=powershell-6
            [psobject]$responseHeaders = @()
            [psobject]$o1 = Invoke-RestMethod -Uri $resturl -Method Get -ResponseHeadersVariable $responseHeaders -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
            [string]$continuationToken = $responseHeaders['X-MS-ContinuationToken']

            Currently the code is based on PowerShell 5.1. Hence why it uses Invoke-WebRequest.

        .Outputs
            Returns a PSObject containing Azure DevOps Security Groups
    #>
    [string]$resturl = Get-GraphGroupsUrl
    [psobject]$response = Invoke-WebRequest -Uri $resturl -Method Get -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
    [psobject]$o1 = (ConvertFrom-Json -InputObject $response.Content).value
    [string]$continuationToken = $response.Headers['X-MS-ContinuationToken']
    while (-not [string]::IsNullOrEmpty($continuationToken)) {
        $resturl = Get-GraphGroupsUrl -Token $continuationToken
        $response = Invoke-WebRequest -Uri $resturl -Method Get -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
        [psobject]$o2 = (ConvertFrom-Json -InputObject $response.Content).value
        $continuationToken = $response.Headers['X-MS-ContinuationToken']
        [psobject]$o3 = $o1
        $o3 += $o2
        $o1 = $o3.psobject.copy()
    }
    $o1
}

function Set-TaskGroupPermissions {
    <#
        .Synopsis
            Set the permissions of a specific task group

        .Outputs
            Returns a PSObject with keys for:
            Success, true if the access control list was updated, otherwise false

        .Parameter Project
            The team project name

        .Parameter TaskgroupName
            The name of the task group for which the permissions will be set

        .Parameter GroupName
            The Azure DevOps group. This can be Contributors, Project Administrators, Release Administrators, Build Administrators

        .Parameter Action
            The security namespace action for the MetaTask (task group) to target. This can be Edit, Delete or Administer
    #>
    param (
        [Parameter(Mandatory = $true)][string]$Project,
        [Parameter(Mandatory = $true)][string]$TaskgroupName,
        [Parameter(Mandatory = $false)][string][ValidateSet('Contributors', 'Project Administrators', 'Release Administrators', 'Build Administrators')]$GroupName = "Contributors",
        [Parameter(Mandatory = $false)][string][ValidateSet('Edit', 'Administer', 'Delete')]$Action = "Edit",
        [Parameter(Mandatory = $false)][string]$RequestSleepTime
    )

    # Validate parameters
    if ([string]::IsNullOrEmpty($Project)) {
        throw "To set task group permissions the team project name is required. It is null or empty"
    }
    if ([string]::IsNullOrEmpty($TaskgroupName)) {
        throw "To set task group permissions the task group name is required. It is null or empty"
    }

    #region group processing
    $o = [PSCustomObject]@{
        SetPermissions = 'NotUpdated'
    }

    [System.Text.RegularExpressions.Regex]$RegEx = [System.Text.RegularExpressions.Regex]::new("^\[$($Project)\]\\$($GroupName)$", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    if ($null -eq $script:Groups) {
        $script:Groups = Get-Groups
    }

    [psobject]$Group = ($script:Groups | Where-Object -FilterScript {$RegEx.IsMatch($_.principalName)})
    if ($Group.PSObject.Properties.Match("principalName").Count) {
        Write-Verbose -Message ("Group [{0}] is valid" -f $Group.principalName) -Verbose
    }
    else {
        throw "The Group [$Project]\$Group is not valid"
    }
    #endregion group processing

    #region project processing
    if ($script:ProjectNameIdLookup.Keys.Count -eq 0) {
        $script:ProjectNameIdLookup = Get-AllAzureDevOpsTeamProjectNames
        if ($script:ProjectNameIdLookup.Keys.Count -eq 0) {
            throw "Unable to retrieve a list of Azure DevOps Team Project Names"
        }
    }
    if ($script:ProjectProperties.Keys.Count -eq 0) {
        $script:ProjectProperties.Add("Name", "")
        $script:ProjectProperties.Add("Id", "")
        $script:ProjectProperties.Add("TaskgroupUrl", "")
    }
    if ($script:ProjectProperties["Name"] -ne $Project) {
        $script:ProjectProperties["Id"] = $script:ProjectNameIdLookup[$Project]
        if ([string]::IsNullOrEmpty($script:ProjectProperties["Id"])) {
            throw "Required id for project [$Project] equates to empty string or null. Expect a project id"
        }
        $script:ProjectProperties["TaskgroupUrl"] = Get-TaskgroupGetUrl -Project $Project
        $script:ProjectProperties["Name"] = $Project
    }
    #endregion project processing

    #region metatask security namespace
    if ($script:MetaTaskSecurityNamespace.count -eq 0) {
        $null = Set-MetaTaskSecurityNamespace
        Write-Verbose -Message ("Found the MetaTask Security Namespace [{0}]" -f $script:MetaTaskSecurityNamespace.displayName)
        if ([psobject]$securityNamespaceAction = ($script:MetaTaskSecurityNamespace.actions | Where-Object -FilterScript {$_.name -eq $Action})) {
            Write-Verbose -Message ("Found security namespace [{0}] [{1}] action with bit value [{2}]" -f $securityNamespaceAction.name, $securityNamespaceAction.displayName, $securityNamespaceAction.bit)
            $script:MetaTaskSecurityNamespaceActionBit = $securityNamespaceAction.bit
        }
        else {
            throw "Security item '$($Action)' not found within security namespace, available values are: $([System.String]::Join(', ', $script:MetaTaskSecurityNamespace.actions.name))"
        }
    }
    #endregion metatask security namespace

    #region set permissions
    if ($script:MetaTaskSecurityNamespace.ContainsKey('actions')) {

        #region ACL processing
        [psobject]$taskgroups = Invoke-RestMethod -Uri $script:ProjectProperties["TaskgroupUrl"] -Method Get -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
        if ($taskgroups.Count -gt 0) {
            if ([psobject]$taskgroup = ($taskgroups.value | Where-Object -FilterScript {$_.name -eq $TaskgroupName})) {
                Write-Verbose -Message ("Found Task group [{0}], and will attempt to update the permissions" -f $taskgroup[0].name) -Verbose
                $resturl = Get-AclUrl -SecurityNamespaceId $script:MetaTaskSecurityNamespace.namespaceId -ProjectId $script:ProjectProperties["Id"] -TaskgroupId $taskgroup[0].id
            }
            else {
                throw "Task group [$TaskgroupName] not found within the Task group(s): $([System.String]::Join(', ', $taskgroups.value.name))"
            }
        }
        else {
            throw "No Task group(s) found within Azure DevOps Project '$($Project)'"
        }

        #region get ACL
        [int]$counter = 0
        [psobject]$accessControlList = @{}
        while ($accessControlList.count -eq 0 -and $counter -ne 4) {
            $accessControlList = Invoke-RestMethod -Uri $resturl -Method Get -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
            if ($accessControlList.count -eq 0) {
                Write-Verbose -Message ("Access Control List for project [{0}] and task group [{1}] empty. Pausing [{2}] seconds before trying again" -f $Project, $TaskgroupName, $RequestSleepTime) -Verbose
                Start-Sleep -Seconds $RequestSleepTime
            }
            $counter++
        }
        #endregion get ACL
        #endregion ACL processing

        #region ACE processing
        if ($accessControlList.count -eq 1) {
            # Get the descriptor of the Security Group
            $resturl = Get-Identity -GroupId $Group.originId
            [psobject]$identity = Invoke-RestMethod -Uri $resturl -Method Get -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
            if ($null -eq ($accessControlList.value[0].acesDictionary | Get-Member -MemberType NoteProperty -Name $identity.descriptor)) {
                # Add security item
                Write-Verbose -Message ("Adding Access Control Entry to [{0}]" -f $accessControlList.value[0].token)
                [pscustomobject]$descriptor = @{
                    "descriptor" = $identity.descriptor
                    "allow"      = $script:MetaTaskSecurityNamespaceActionBit
                    "deny"       = $script:MetaTaskSecurityNamespaceActionBit
                }
                $accessControlList.value[0].acesDictionary | Add-Member -MemberType NoteProperty -Name $identity.descriptor -Value $descriptor
                $resturl = Get-AclUrl -SecurityNamespaceId $script:MetaTaskSecurityNamespace.namespaceId
                [psobject]$jsonContent = $accessControlList | ConvertTo-Json -Depth 10 -Compress
                [psobject]$result = Invoke-WebRequest -Uri $resturl -Method Post -Body $jsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo) }
                if ($result.StatusCode -ne 204) {
                    throw "Failed to add permission. StatusCode '$($result.StatusCode)', StatusDescription '$($result.StatusDescription)'"
                }
                else {
                    $o.SetPermissions = 'Updated'
                }
            }
            else {
                # Validate the permissions
                if ($accessControlList.value[0].acesDictionary.$($identity.descriptor).deny -ne $script:MetaTaskSecurityNamespaceActionBit) {
                    Write-Verbose -Message ("Change permission from [{0}] to [{1}]" -f $accessControlList.value[0].acesDictionary.$($identity.descriptor).deny, $script:MetaTaskSecurityNamespaceActionBit) -Verbose
                    $accessControlList.value[0].acesDictionary.$($identity.descriptor).deny = $script:MetaTaskSecurityNamespaceActionBit
                    $resturl = Get-AclUrl -SecurityNamespaceId $script:MetaTaskSecurityNamespace.namespaceId
                    [psobject]$jsonContent = $accessControlList | ConvertTo-Json -Depth 10 -Compress
                    [psobject]$result = Invoke-WebRequest -Uri $resturl -Method Post -Body $jsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $script:AuthorisationType, $script:Base64AuthInfo); ContentType = ("application/json"); }
                    if ($result.StatusCode -ne 204) {
                        throw "Failed to add permission. StatusCode '$($result.StatusCode)', StatusDescription '$($result.StatusDescription)'"
                    }
                    else {
                        $o.SetPermissions = 'Updated'
                    }
                }
                else {
                    Write-Verbose -Message ("No changes required for the Access Control Entry") -Verbose
                }
            }
        }
        else {
            throw "Found $($accessControlList.count) number of ACL entries. Expecting just 1"
        }
        #endregion ACE processing
    }
    else {
        throw "Security namespace 'MetaTask' not found"
    }    
    #endregion set permissions
    $o
}
