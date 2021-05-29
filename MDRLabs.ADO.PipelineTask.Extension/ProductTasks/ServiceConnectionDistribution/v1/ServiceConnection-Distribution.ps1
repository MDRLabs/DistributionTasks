<#
     .Synopsis
        Add or update an Azure DevOps Service Connection. 

     .Description
        The script calls the serviceconnection REST API to create a new, or update an existing Azure DevOps
        service connection. The scipt is packaged in as an Azure DevOps Build Pipeline extension.

    .Parameter Token
    .Parameter Pat
        The user's Personal Access Token (PAT) with the correct scopes for creating and updating Serviceconnection.

    .Parameter CommitMessage
    .Parameter Comment
        The reason for the change. This would normally be the last commit message

    .Parameter AzureDevOpsTeamProjects
    .Parameter Projects
        Projects are the Azure DevOps Team Project(s) to target. If no value is specified, 
        all Azure DevOps Team Projects will be targeted.
        If providing multiple teams, use a comma (,) to separate each and quote the names if they
        contain spaces.

    .Parameter ExcludeTeamProjects
    .Parameter Exclude
        To exclude one or more Azure DevOps Team Projects from the run use this parameter.
        If providing multiple teams, use a comma (,) to separate each and quote the names if they
        contain spaces.

    .Parameter Folder
        This is the name of a single directory. Unless the value contains a fully qualified directory path, the 
        directory is assumed to be relative to the script folder. Only JSON files in this folder will
        be included in the run.

    .Parameter Jsonfiles
    .Parameter Files
        One or more JSON files containing a valid service connection definition. 
        If the files are not in the same directory as the script they must be fully qualified.
        If providing multiple files, use a comma (,) to separate each.

    .Inputs
        See Jsonfiles and Folder parameters

    .Outputs
        Exitcode 0 if successful, otherwise 1

    .Example
        Run the script from a PowerShell CLI
          Maintain-Pipelines [-verbose] -Token pat [-Projects "teamproject1","teamproject2" ... | -Exclude "teamproject1","teamproject2" ...] -Folder directoryname | -Files "file1","file2","file3" ...
        Run the script an from Azure DevOps PowerShell task
          Maintain-Pipelines [-debug] [-verbose] -Token pat [-Projects "teamproject1","teamproject2" ... | -Exclude "teamproject1","teamproject2" ...] -Folder directoryname | -Files "file1","file2","file3" ...

    .Link
        https://martin77s.wordpress.com/2014/06/17/powershell-scripting-best-practices/
        https://docs.microsoft.com/en-us/rest/api/azure/devops/?view=azure-devops-rest-5.1
        https://docs.microsoft.com/en-us/azure/devops/integrate/concepts/integration-bestpractices?view=vsts
        https://docs.microsoft.com/en-us/azure/devops/integrate/concepts/rate-limits?view=azure-devops&amp%3Btabs=new-nav&viewFallbackFrom=vsts&tabs=new-nav
        https://docs.microsoft.com/en-us/azure/devops/pipelines/library/service-endpoints?view=azure-devops&tabs=yaml
.Notes
        1. If running the script as an Azure DevOps extension the -Token variable should be stored as a secret variable.
        2. The -Projects and the -Exclude parameters are mutually exclusive. If the run contains a list of Team 
           Project names it doesn't make sense to also specify an Exclude list. Only include the latter if no
           -Project parameter is specified and there is a need to exclude ome or more team projects from the run.
        3. The -Folder and -Files parameters are mutually exclusive. If specified the -Folder parameter must 
           resolve to a valid directory containing JSON files, otherwise the the files specified with -Files
           will be used.
        4. RATE LIMITS
           Refer to the rate-limits wiki (link is included under links)
#>

[CmdletBinding(DefaultParameterSetName = "Folder")]
Param(
    # PAT / Token parameter
    [Parameter(Mandatory = $true, HelpMessage = 'Your Azure DevOps PAT with all scopes')]
    [Alias('Pat')]
    [string]
    $Token,

    # Commit Message parameter
    [Parameter(Mandatory = $true, HelpMessage = 'Last Commit message or reason for change')]
    [Alias('Comment')]
    [string]
    $CommitMessage,

    # Projects / AzureDevops Team Projects parameter
    [Parameter(Mandatory = $true, ParameterSetName = 'Projects', HelpMessage = 'A comma/quoted list of one or more team project names')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Folder')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Files')]
    [AllowEmptyCollection()]
    [Alias('Projects')]
    [string[]]
    $AzureDevOpsTeamProjects,

    # Folder parameter
    [Parameter(Mandatory = $true, ParameterSetName = 'Folder', HelpMessage = 'The name of a folder containing Service connection')]
    [string]
    $Folder,

    # Exclude / Excludeteamprojects parameter
    [Parameter(Mandatory = $true, ParameterSetName = 'Exclude', HelpMessage = 'A comma/quoted list of one or more team project names to ignore')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Folder')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Files')]
    [AllowEmptyCollection()]
    [Alias('Exclude')]
    [string[]]
    $ExcludeTeamProjects,

    # Files / JsonFiles parameter
    [Parameter(Mandatory = $true, ParameterSetName = 'Files', HelpMessage = 'A comma/quoted list of JSON Service connection files')]
    [AllowEmptyCollection()]
    [Alias('Files')]
    [string[]]
    $Jsonfiles
)

#******************************************************************************************************************
# Script body
# Execution begins here
#******************************************************************************************************************
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop';

try {
    $global:ScriptFolder = Split-Path -Parent $myInvocation.MyCommand.Path

    Write-Verbose -Message ("Serviceconnection-Distribution run has started") -Verbose
    Write-Verbose -Message ("The PowerShell version is [{0}]" -f $PSVersionTable.PSVersion)
    Write-Verbose -Message ("The chosen parameter set name is [{0}]" -f $PSCmdlet.ParameterSetName)

    [int]$Exitcode = 0
    [psobject]$JsonContent = @{}
    [System.Collections.Generic.Dictionary[string, string]]$ProjectDictionary = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    [System.Collections.Generic.List[string]]$Projects = New-Object 'System.Collections.Generic.List[string]'
    [System.Collections.Generic.List[string]]$ProjectExclusions = New-Object 'System.Collections.Generic.List[string]'
    [System.Collections.Generic.List[string]]$Files = New-Object 'System.Collections.Generic.List[string]'

    # Expects the DesAzure.PowerShell.Vsts.RestApi.Helper module to be imported.
    if (-not(Get-Module -name "DesAzure.Ps.Azure.DevOps.RestApi.Helper")) {
        $localmodules = Join-Path $ScriptFolder "Modules\DesAzure.Ps.Azure.DevOps.RestApi.Helper.psm1"
        import-module $localmodules -Force -NoClobber -Global -ErrorAction Stop
        if ($null -eq (Get-Module -Name "DesAzure.Ps.Azure.DevOps.RestApi.Helper").Name) {
            throw "Module DesAzure.Ps.Azure.DevOps.RestApi.Helper not loaded. Install-Module DesAzure.Ps.Azure.DevOps.RestApi.Helper and re-run"
        }
    }

    # Set/Get Azure DevOps authorisation token
    Set-PersonalAccessToken -Token $Token
    [string]$AuthorisationType = Get-AuthorisationType
    [string]$EncodedToken = Get-Base64AuthInfo 

    # Create the dictionary of project names and ids
    $ProjectDictionary = Get-AllAzureDevOpsTeamProjectNames

    # Determine which team projects to include in this run
    if ([string]::IsNullOrEmpty($AzureDevOpsTeamProjects)) {
        $Projects = $ProjectDictionary.Keys
        Write-Verbose -Message ("All team projects will be targeted. Total teams is [{0}]" -f $Projects.Count) -Verbose
    }
    else {
        foreach ($p in $AzureDevOpsTeamProjects) {
            if ($ProjectDictionary.ContainsKey($p)) {
                $Projects.Add($p)
            }
            else {
                Write-Verbose -Message ("Ignoring project [$p] because it does not exist in this organisation") -Verbose
            }
        }
        Write-Verbose -Message ("Run applies to [{0}] projects. These are [{1}]" -f $Projects.Count, $([System.String]::Join(', ', ($Projects | Sort-Object)))) -Verbose
    }
    if (-not [string]::IsNullOrEmpty($ExcludeTeamProjects) -and $ExcludeTeamProjects.Count -gt 0) {
        foreach ($p in $ExcludeTeamProjects) {
            if ($ProjectDictionary.ContainsKey($p)) {
                $ProjectExclusions.Add($p)
            }
            else {
                Write-Verbose -Message ("Ignoring exclude project [$p] because it does not exist in this organisation") -Verbose
            }
        }
        if ($ProjectExclusions.Count -eq 0) {
            throw "The specified Project Exclusion List contains no valid Team Projects. Terminating to prevent targetting all Team Projects"
        }
        Write-Verbose -Message ("All teams projects will be targeted except [{0}]" -f $([System.String]::Join(', ', ($ProjectExclusions | Sort-Object)))) -Verbose
    }

    # Determine which [JSON] files will be included in this run. If the -Folder
    # parameter is specified use it, unless it is invalid and/or doesn't contain
    # JSON files, in which case fall back to the -Files parameter.
    if (-not [string]::IsNullOrEmpty($Folder) -and (Test-Path -Path $Folder)) {
        [psobject]$f = Get-ChildItem -Path "$Folder\*" -Include "*.json"
        foreach ($element in $f) {
            $Files.Add($element.FullName)
        }
    }
    else {
        $Files.AddRange($Jsonfiles)
    }

    # If there are no files to process stop the run
    if ($Files.Count -eq 0) {
        Write-Warning -Message "The current run has no valid JSON files (service connections) to distribute. Terminating the run"
    }
    else {  
        Write-Verbose -Message ("There are [{0}] JSON files (Service connections) to process. They are [{1}]" -f $Files.Count, $([System.String]::Join(', ', ($Files | Sort-Object))))

        # Identify the security-namespace for MetaTask
        foreach ($project in $Projects) {

            if ($ProjectExclusions.Contains($project)) {
                Write-Verbose -Message "Excluding the current project [$project] as requested" -Verbose
                continue
            }
            
            Write-Verbose -Message "The current team project name is [$project]" -Verbose
    
            foreach ($jsonFile in $Files) {

                Write-Verbose -Message "The current file is [$jsonFile]" -Verbose
                [PSCustomObject]$jsonObject = Get-JsonContent -Filename $jsonFile -DefinitionType "serviceconnection"
        
                if ($jsonObject.ValidJson) {
                    
                    [string]$jsonObjectName = $jsonObject.Name
                    [psobject]$jsonServiceConnectionObject = $jsonObject.JsonObject
                    [psobject]$jsonContent = $jsonServiceConnectionObject.serviceconnection[0].properties | ConvertTo-Json -Compress

                    $endpointName = $jsonServiceConnectionObject.serviceconnection[0].properties.name
                    [string]$apiGetUri = Get-ServiceConnectionGetByNameUrl -Project $project -EndpointName $endpointName
                    [psobject]$vstsServiceConnection = Invoke-RestMethod -Uri $apiGetUri -Method GET -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken); ContentType = ("application/json"); }                    
                    if ($vstsServiceConnection.count -eq 0)
                    {
                        # Create new service connection
                        Write-Verbose -Message "Service connection endpoint doesn't exist. Creating [$jsonObjectName]"
                        [string]$ApiPostUri = Get-ServiceConnectionPostUrl -Project $project
                        Invoke-RestMethod -Uri $ApiPostUri -Method POST -Body $jsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) } | Out-Null
                        Write-Verbose -Message "Service connection [$jsonObjectName] for Team project [$project] has been created"
                    }
                    else {
                        # Update service connection
                        Write-Verbose -Message "Service connection endpoint already exist. Updating [$jsonObjectName]"
                        $endpointId = $vstsServiceConnection.value[0].id
                        [string]$ApiPutUri = Get-ServiceConnectionPutUrl -Project $project -EndpointId $endpointId
                        Invoke-RestMethod -Uri $ApiPutUri -Method PUT -Body $jsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) } | Out-Null
                        Write-Verbose -Message "Service connection [$jsonObjectName] for Team project [$project] has been updated"
                    }
                }
                else {
                    Write-Warning -Message "Input file [$jsonFile] does not contain a valid service connection definition. Ignoring the file"
                }
            }   
        }
    }
}
catch [Exception] {
    Write-Error -Message ("Caught an exception of type [{0}], and reason [{1}]" -f $_.Exception.GetType().FullName, $_.Exception.Message) -ErrorID "1" -Targetobject $_
    if($null -ne $_.ErrorDetails){
        Write-Error -Message ("Caught an exception of type [{0}], and reason [{1}]" -f $_.Exception.GetType().FullName, $_.ErrorDetails.Message) -ErrorID "1" -Targetobject $_
    }
    $Exitcode = 1
}
finally {
    Write-Verbose "Maintain Pipelines service connection execution ended" -Verbose
}

Set-StrictMode -Off
Exit $Exitcode
