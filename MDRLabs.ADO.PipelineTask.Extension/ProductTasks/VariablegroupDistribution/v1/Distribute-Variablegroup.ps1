<#
     .SYNOPSIS
        Add or update a Vsts Variable Group

     .DESCRIPTION
        The script calls the variablegroups REST API to create a new, or update an existing Vsts
        Variable Group. The script requires the CredentialManager module to be installed.
        If necessary first run Install-Module -Name CredentialManager. See LINKS.

        The script will prompt for a valid VSTS Personal Access Token which permissions to
        create Build resources. If necessary create a PAT with all scopes.
        Valid credentials will be stored by Windows Credential Manager and reused when the
        script is re-run.

    .PARAMETER Token
    .PARAMETER Pat
        The user's Personal Access Token (PAT) with the correct scopes for creating and updating Variablegroups.
        This must be a SecureString.

    .PARAMETER AzureDevOpsTeamProjects
    .PARAMETER Projects
        Projects are the Azure DevOps Team Project(s) to target. If no value is specified, 
        all Azure DevOps Team Projects will be targeted.
        If providing multiple teams, use a comma (,) to separate each and quote the names if they
        contain spaces.

    .PARAMETER ExcludeTeamProjects
    .PARAMETER Exclude
        To exclude one or more Azure DevOps Team Projects from the run use this parameter.
        If providing multiple teams, use a comma (,) to separate each and quote the names if they
        contain spaces.

    .PARAMETER Folder
        This is the name of a single directory. Unless the value contains a fully qualified directory path, the 
        directory is assumed to be relative to the script folder. Only JSON files in this folder will
        be included in the run.

    .PARAMETER Jsonfiles
    .PARAMETER Files
        One or more JSON files containing a valid variablegroup definition. 
        If the files are not in the same directory as the script they must be fully qualified.
        If providing multiple files, use a comma (,) to separate each.

    .INPUTS
        See Jsonfiles parameter

    .OUTPUTS
        Exitcode 0 if successful, otherwise 1

    .EXAMPLE
        Run the script from a PowerShell CLI
        Distribute-Variablegroup [-verbose] -Token pat [-Projects "teamproject1","teamproject2" ... | -Exclude "teamproject1","teamproject2" ...] -Folder directoryname | -jsonfiles "file1","file2","file3" ...
        
        Run the script an from Azure DevOps PowerShell variablegroups
        Distribute-Variablegroup [-debug] [-verbose] -Token pat [-Projects "teamproject1","teamproject2" ... | -Exclude "teamproject1","teamproject2" ...] -Folder directoryname | -jsonfiles "file1","file2","file3" ...

    .LINK
        https://docs.microsoft.com/en-us/rest/api/azure/devops/?view=azure-devops-rest-5.1
        https://docs.microsoft.com/en-us/azure/devops/integrate/concepts/integration-bestpractices?view=vsts
        https://docs.microsoft.com/en-us/rest/api/vsts/distributedtask/variablegroups
        https://github.com/Microsoft/azure-pipelines-task-lib/blob/master/tasks.schema.json
        https://martin77s.wordpress.com/2014/06/17/powershell-scripting-best-practices/

    .NOTES
        Invoke the VSTS distributetask variablegroups REST Api, passing a valid JSON object containing a Vsts Variablegroup.
        If you are providing your own JSON file, not one from the Standard Build Definitions, make sure you provide a
        comment property with a description of the change. Leave the value of the id property as 1.
#>

[CmdletBinding()]
Param(   
    # PAT / Token parameter
    [Parameter(Mandatory = $true, HelpMessage = 'Your Azure DevOps PAT with all scopes')]
    [Alias('Pat')]
    [string]
    $Token,

    # Projects / AzureDevops Team Projects parameter
    [Parameter(Mandatory = $true, ParameterSetName = 'Projects', HelpMessage = 'A comma/quoted list of one or more team project names')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Folder')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Files')]
    [AllowEmptyCollection()]
    [Alias('Projects')]
    [string[]]
    $AzureDevOpsTeamProjects,

    # Folder parameter
    [Parameter(Mandatory = $true, ParameterSetName = 'Folder', HelpMessage = 'The name of a folder containing Variablegroup')]
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
    [Parameter(Mandatory = $true, ParameterSetName = 'Files', HelpMessage = 'A comma/quoted list of JSON Variablegroup files')]
    [AllowEmptyCollection()]
    [Alias('Files')]
    [string[]]$jsonfiles
)

#******************************************************************************************************************
# Script body
# Execution begins here
#******************************************************************************************************************
Set-StrictMode -Version Latest

try {
    $global:ScriptFolder = Split-Path -Parent $myInvocation.MyCommand.Path

    Write-Verbose -Message ("Variable Group Distribution-Pipelines run has started") -Verbose
    Write-Verbose -Message ("The PowerShell version is [{0}]" -f $PSVersionTable.PSVersion)

    [int]$Exitcode = 0
    [psobject]$JsonContent = @{}
    [System.Collections.Generic.Dictionary[string, string]]$ProjectDictionary = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    [System.Collections.Generic.List[string]]$Projects = New-Object 'System.Collections.Generic.List[string]'
    [System.Collections.Generic.List[string]]$ProjectExclusions = New-Object 'System.Collections.Generic.List[string]'
    [System.Collections.Generic.List[string]]$Files = New-Object 'System.Collections.Generic.List[string]'

    # Expects the MDRLabs.PowerShell.Vsts.RestApi.Helper module to be imported.
    if (-not(Get-Module -name "MDRLabs.Ps.Ado.RestApi.HelperMDRLabs.ADORestApiHelper")) {
        $localmodules = Join-Path $ScriptFolder "Modules\MDRLabs.Ps.Ado.RestApi.HelperMDRLabs.ADORestApiHelper.psm1"
        import-module $localmodules -Verbose -Force -NoClobber -Global -ErrorAction Stop
       
        if ($null -eq (Get-Module -Name "MDRLabs.Ps.Ado.RestApi.HelperMDRLabs.ADORestApiHelper").Name) {
            throw "Module MDRLabs.Ps.Ado.RestApi.HelperMDRLabs.ADORestApiHelper not loaded. Install-Module MDRLabs.Ps.Ado.RestApi.HelperMDRLabs.ADORestApiHelper and re-run"
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

    # Determine which team projects to exclude in this run
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
        Write-Verbose -Message "The current run has no valid JSON files (variablegroups) to distribute. Terminating the run" -Verbose
    }
    else {  
        Write-Verbose -Message ("There are [{0}] JSON files (variablegroups) to process. They are [{1}]" -f $Files.Count, $([System.String]::Join(', ', ($Files | Sort-Object))))

        foreach ($project in $Projects) {

            if ($ProjectExclusions.Contains($project)) {
                Write-Verbose -Message "Excluding the current project [$project] as requested" -Verbose
                continue
            }

            Write-Verbose -Message "The current team project name is [$project]" -Verbose

            # Read and validate each JSON file and either Add, Update or ignore the Variablegroup
            foreach ($jsonFile in $Files) {

                [PSCustomObject]$jsonObject = Get-JsonContent -Filename $jsonFile -PropertyName "variables"

                if ($jsonObject.ValidJson) {
                    [string]$apiGetUri = Get-VariablegroupGetUrl -Project $project -GroupName $jsonObject.Name
                    
                    #Get all Variable Group
                    [psobject]$vstsVariableGroup = Invoke-RestMethod -Uri $apiGetUri -Method GET -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken); ContentType = ("application/json"); }
                    [psobject]$jsonContent = $jsonObject.JsonContent
                    [string]$jsonObjectName = $jsonObject.Name

                    if ($vstsVariableGroup.count -eq 0) {
                        
                        #Create Variable Group
                        Write-Verbose -Message "Variablegroup doesn't exist. Creating [$jsonObjectName]"
                        [string]$ApiPostUri = Get-VariablegroupPostUrl -Project $project
                        Invoke-RestMethod -Uri $ApiPostUri -Method POST -Body $jsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) } | Out-Null
                        Write-Verbose -Message "Variable Group [$jsonObjectName] for Team project [$project] has been created"
                    }
                    else {
                        [string]$vgid = $vstsVariableGroup.value[0].id
                        [psobject]$variableGroupObject = $jsonObject.JsonObject

                        Write-Verbose -Message "Found Variablegroup [$jsonObjectName] with Id [$vgid]. Updating it."
                        
                        if ($null -eq $variableGroupObject.id) {
                            Add-Member -inputobject $variableGroupObject -membertype NoteProperty -Name 'id' -Value $vgid
                        }
                        else {
                            $variableGroupObject.id = $vgid
                        }
                        
                        [string]$apiUpdateUri =  Get-VariablegroupPutUrl -Project $project -VariableGroupObjectId $variableGroupObject.id
                        $updatedJsonContent = $variableGroupObject | ConvertTo-Json -Depth 3

                        #Update Variable Group
                        Invoke-RestMethod -Uri $apiUpdateUri -Method PUT -Body $updatedJsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) } | Out-Null
                        Write-Verbose -Message "Variable Group [$jsonObjectName] for Team project [$project] has been updated"
                    }
                }
                else {
                    Write-Warning -Message "Input file [$jsonFile] does not contain a valid Variablegroup definition. Ignoring the file"
                }
            }
        }
    }
}
catch [Exception] {
    Write-Error -Message ("Caught an exception of type [{0}], and reason [{1}]" -f $_.Exception.GetType().FullName, $_.Exception.Message) -ErrorID "1" -Targetobject $_
    $Exitcode = 1
}
finally {
    Write-Verbose "Maintain Pipelines distribute Variablegroup execution ended" -Verbose
}

Set-StrictMode -Off
Exit $Exitcode