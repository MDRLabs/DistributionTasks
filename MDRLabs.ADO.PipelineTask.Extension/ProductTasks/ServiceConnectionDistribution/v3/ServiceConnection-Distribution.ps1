<#
     .Synopsis
        Add or update an Azure DevOps Service Connection. 

     .Description
        The script calls the serviceconnection REST API to create a new, or update an existing Azure DevOps
        service connection. The scipt is packaged in as an Azure DevOps Build Pipeline extension.

    .PARAMETER IsAdoCloud
    .PARAMETER IsAdoCloud
        The Azure DevOps instance to tearget i.e. On-Premises or Cloud 

    .PARAMETER AdoServerUrl
    .PARAMETER AdoServerUrl
        Azure DevOps server url based on selection of IsAdoCloud

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
    # IsAdoCloud parameter
    [Parameter(Mandatory = $true, HelpMessage = 'Your Azure DevOps Cloud Instance')]
    [bool]
    $IsAdoCloud=1,

    # ServerUrl parameter
    [Parameter(Mandatory = $false)]
    [string]
    $ServerUrl='',
        
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
    $ErrorBuilder = [System.Text.StringBuilder]::new()

    # Execution summary arrays
    [System.Collections.Generic.List[string]]$FilesSkipped = New-Object 'System.Collections.Generic.List[string]'
    $FailedProjectCollection = @{}
    $SuccededProjectCollection = @{}

    # Expects the DesAzure.PowerShell.Vsts.RestApi.Helper module to be imported.
    if (-not(Get-Module -name "ADORestApiHelper")) {
        $localmodules = Join-Path $ScriptFolder "Modules\ADORestApiHelper.psm1"
        import-module $localmodules -Force -NoClobber -Global -ErrorAction Stop
        if ($null -eq (Get-Module -Name "ADORestApiHelper").Name) {
            throw "Module ADORestApiHelper not loaded. Install-Module ADORestApiHelper and re-run"
        }
    }

    # Set/Get Azure DevOps authorisation token
    Set-PersonalAccessToken -Token $Token
    [string]$AuthorisationType = Get-AuthorisationType
    [string]$EncodedToken = Get-Base64AuthInfo 

    if($IsAdoCloud)
    {
        $ServerUrl = Get-CloudOrganizationUrl
    }

    # Create the dictionary of project names and ids
    $ProjectDictionary = Get-AllAzureDevOpsTeamProjectNames -AdoServerUrl $ServerUrl

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

        [System.Collections.Generic.List[string]]$FileProjectFailed = New-Object 'System.Collections.Generic.List[string]'
        [System.Collections.Generic.List[string]]$FileProjectPassed = New-Object 'System.Collections.Generic.List[string]'

        # Identify the security-namespace for MetaTask
        foreach ($jsonFile in $Files) {
            Write-Verbose -Message "The current file is [$jsonFile]" -Verbose

            try {
                [PSCustomObject]$jsonObject = Get-JsonContent -Filename $jsonFile -DefinitionType "serviceconnection"
                if ($jsonObject.ValidJson) {
                
                    [psobject]$jsonServiceConnectionObject = $jsonObject.JsonObject
                    [psobject]$jsonContent = $jsonServiceConnectionObject.serviceconnection[0].properties | ConvertTo-Json -Compress
                    $endpointName = $jsonServiceConnectionObject.serviceconnection[0].properties.name

                    # Identify the security-namespace for MetaTask
                    foreach ($project in $Projects) {
                                            
                        if ($ProjectExclusions.Contains($project)) {
                            Write-Verbose -Message "Excluding the current project [$project] as requested" -Verbose
                            continue
                        }

                        Write-Verbose -Message "The current team project name is [$project]" -Verbose

                        [string]$apiGetUri = Get-ServiceConnectionGetByNameUrl -AdoServerUrl $ServerUrl -Project $project -EndpointName $endpointName
                        [psobject]$vstsServiceConnection = Invoke-RestMethod -Uri $apiGetUri -Method GET -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken); ContentType = ("application/json"); }                    
                        if ($vstsServiceConnection.count -eq 0)
                        {
                            try {
                                # Create new service connection
                                [string]$ApiPostUri = Get-ServiceConnectionPostUrl -AdoServerUrl $ServerUrl -Project $project
                                Invoke-RestMethod -Uri $ApiPostUri -Method POST -Body $jsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) } | Out-Null
                                $FileProjectPassed.Add($project)
                            }
                            catch [Exception]{
                                $FileProjectFailed.Add($project)
                                if($null -ne $_.ErrorDetails){
                                    $errorJson = Get-ValidateJsonSchema -JsonText $_.ErrorDetails.Message
                                    if($null -ne $errorJson)
                                    {
                                        [void]$ErrorBuilder.Append([string]::Format("Project [{0}], File [{1}], Reason [{2}] `r`n", $project, [System.IO.Path]::GetFileName($jsonFile), $errorJson.JsonObject.message))
                                    }
                                    else {
                                        [void]$ErrorBuilder.Append([string]::Format("Project [{0}], File [{1}], Reason [{2}] `r`n", $project, [System.IO.Path]::GetFileName($jsonFile), $_.ErrorDetails.Message))
                                    }
                                }
                                else {
                                    [void]$ErrorBuilder.Append([string]::Format("Project [{0}], File [{1}], Reason [{2}] `r`n", $project, [System.IO.Path]::GetFileName($jsonFile), $_.Exception.Message))
                                }
                                Continue 
                            }
                        }
                        else {
                            try {
                                # Update service connection
                                $endpointId = $vstsServiceConnection.value[0].id
                                [string]$ApiPutUri = Get-ServiceConnectionPutUrl -AdoServerUrl $ServerUrl -Project $project -EndpointId $endpointId
                                Invoke-RestMethod -Uri $ApiPutUri -Method PUT -Body $jsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) } | Out-Null
                                $FileProjectPassed.Add($project)
                            }
                            catch {
                                $FileProjectFailed.Add($project)
                                if($null -ne $_.ErrorDetails){
                                    $errorJson = Get-ValidateJsonSchema -JsonText $_.ErrorDetails.Message
                                    if($null -ne $errorJson)
                                    {
                                        [void]$ErrorBuilder.Append([string]::Format("Project [{0}], File [{1}], Reason [{2}] `r`n", $project, [System.IO.Path]::GetFileName($jsonFile), $errorJson.JsonObject.message))
                                    }
                                    else {
                                        [void]$ErrorBuilder.Append([string]::Format("Project [{0}], File [{1}], Reason [{2}] `r`n", $project, [System.IO.Path]::GetFileName($jsonFile), $_.ErrorDetails.Message))
                                    }
                                }
                                else {
                                    [void]$ErrorBuilder.Append([string]::Format("Project [{0}], File [{1}], Reason [{2}] `r`n", $project, [System.IO.Path]::GetFileName($jsonFile), $_.Exception.Message))
                                }
                                Continue 
                            }
                        }
                    }    
                }
                else {
                    $FilesSkipped.Add($jsonFile)
                }

                if($FileProjectPassed.Count -gt 0)
                {
                    $SuccededProjectCollection.Add($jsonFile, $([System.String]::Join(', ', ($FileProjectPassed | Sort-Object))))
                }
                
                if($FileProjectFailed.Count -gt 0)
                {
                    $FailedProjectCollection.Add($jsonFile, $([System.String]::Join(', ', ($FileProjectFailed | Sort-Object))))
                }

                $FileProjectFailed.Clear()
                $FileProjectPassed.Clear()
            }
            catch [Exception]{
                $FilesSkipped.Add($jsonFile)
                if($null -ne $_.ErrorDetails){
                    $errorJson = Get-ValidateJsonSchema -JsonText $_.ErrorDetails.Message
                    if($null -ne $errorJson){
                        [void]$ErrorBuilder.Append([string]::Format("File [{0}], Reason [{1}] `r`n", [System.IO.Path]::GetFileName($jsonFile), $errorJson.JsonObject.message))
                    }
                    else {
                        [void]$ErrorBuilder.Append([string]::Format("File [{0}], Reason [{1}] `r`n", [System.IO.Path]::GetFileName($jsonFile), $_.ErrorDetails.Message))
                    }
                }
                else {
                    [void]$ErrorBuilder.Append([string]::Format("File [{0}], Reason [{1}] `r`n", [System.IO.Path]::GetFileName($jsonFile), $_.Exception.Message)) 
                }
                Continue
            }
        }
        
        $SummaryBuilder = [System.Text.StringBuilder]::new()

        [void]$SummaryBuilder.Append("------------------------------------------------`n")
        [void]$SummaryBuilder.Append("    ServiceConnection distribution summary report   `n")
        [void]$SummaryBuilder.Append("------------------------------------------------`n")
        [void]$SummaryBuilder.Append("`n")

        # Projects found in organization
        $organizationName = Get-OrganizationName -IsAdoCloud $IsAdoCloud -AdoServerUrl $ServerUrl
        if (-not [string]::IsNullOrEmpty($Projects) -and $Projects.Count -gt 0) {
            [void]$SummaryBuilder.Append([string]::Format("`nRun completed in [{0}] projects: [{1}]", $Projects.Count, $([System.String]::Join(', ', ($Projects | Sort-Object)))))
        }
        else {
            [void]$SummaryBuilder.Append("`n Project(s) not found in organization '$organizationName'")
        }
        [void]$SummaryBuilder.Append("`n")

        # Input files
        [void]$SummaryBuilder.Append([string]::Format("`n[{0}] JSON files (Serviceconnections) were processed. They are`n", $Files.Count))
        ForEach ($fileName in $Files){ 
            [void]$SummaryBuilder.Append([string]::Format("{0}",[System.IO.Path]::GetFileName($fileName)))
            [void]$SummaryBuilder.Append("`n")
        }

        # Serviceconnections distribution execution summary dashboard: Invalid input files        
        if (-not [string]::IsNullOrEmpty($FilesSkipped) -and $FilesSkipped.Count -gt 0) {
            $SkippedFileCount = $FilesSkipped.Count
            [void]$SummaryBuilder.Append("`n Following input files($SkippedFileCount) were ignored due to invalid Serviceconnections definition.`n")
            foreach ($fileskipped in $FilesSkipped) {
                [void]$SummaryBuilder.Append([string]::Format("{0}",[System.IO.Path]::GetFileName($fileskipped)))
                [void]$SummaryBuilder.Append("`n")
            }
        }

        # Serviceconnections distribution execution summary dashboard: successful projects
        if (-not [string]::IsNullOrEmpty($SuccededProjectCollection) -and $SuccededProjectCollection.Count -gt 0) {
            [void]$SummaryBuilder.Append("`nServiceconnections create/update successfully for input file: [projects] `n")
            ForEach ($itemKey in $SuccededProjectCollection.Keys){ 
                $itemValue = '[' + $SuccededProjectCollection[$itemKey] + ']'
                [void]$SummaryBuilder.Append([string]::Format("{0} : {1}",[System.IO.Path]::GetFileName($itemKey), $itemValue))
                [void]$SummaryBuilder.Append("`n")
            }
            [void]$SummaryBuilder.Append("`n")
        }

        # Serviceconnections distribution execution summary dashboard: failed projects
        if (-not [string]::IsNullOrEmpty($FailedProjectCollection) -and $FailedProjectCollection.Count -gt 0) {
            [void]$SummaryBuilder.Append("`nServiceconnections create/update failed for input file: [projects]`n")
            ForEach ($itemKey in $FailedProjectCollection.Keys){ 
                $itemValue = '[' + $FailedProjectCollection[$itemKey] + ']'
                [void]$SummaryBuilder.Append([string]::Format("{0} : {1}",[System.IO.Path]::GetFileName($itemKey), $itemValue))
                [void]$SummaryBuilder.Append("`n")
            }
            [void]$SummaryBuilder.Append("`n")
        }

        # Display all error message text
        if(-not [string]::IsNullOrEmpty($ErrorBuilder.ToString()))
        {
            # Mark pipeline as failure and dissplay all errors 
            [void]$SummaryBuilder.Append("`nOne or more projects failed while distributing Serviceconnections. Details of errors can be found below `n")
            [void]$SummaryBuilder.Append($ErrorBuilder.ToString())
            [void]$SummaryBuilder.Append("`n")
        }

        [void]$SummaryBuilder.Append("------------------------------------------------`n")
        $SummaryBuilder.ToString()
    }
}
catch [Exception] {
    if($null -ne $_.ErrorDetails){
        $errorJson = Get-ValidateJsonSchema -JsonText $_.ErrorDetails.Message
        if($null -ne $errorJson){
            [void]$ErrorBuilder.Append([string]::Format("Error encountered during the execution Type [{0}], reason [{1}] `r`n", $_.Exception.GetType().FullName, $errorJson.JsonObject.message))
        }
        else {
            [void]$ErrorBuilder.Append([string]::Format("Error encountered during the execution Type [{0}], reason [{1}] `r`n", $_.Exception.GetType().FullName, $_.ErrorDetails.Message))
        }
    }
    else {
        [void]$ErrorBuilder.Append([string]::Format("Error encountered during the execution Type [{0}], reason [{1}] `r`n", $_.Exception.GetType().FullName, $_.Exception.Message)) 
    }

    $ErrorBuilder.ToString()
}
finally {
    Write-Verbose "Maintain Pipelines Service Connection execution ended" -Verbose
    if (-not [string]::IsNullOrEmpty($FailedProjectCollection) -and $FailedProjectCollection.Count -gt 0){
        throw "One or more error(s) encountered during the execution. Please refer the summary above."
    }
}

Set-StrictMode -Off
Exit $Exitcode
