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
    $ErrorBuilder = [System.Text.StringBuilder]::new()

    # Execution summary arrays
    [System.Collections.Generic.List[string]]$FilesSkipped = New-Object 'System.Collections.Generic.List[string]'
    $FailedProjectCollection = @{}
    $SuccededProjectCollection = @{}

    # Expects the DesAzure.PowerShell.Vsts.RestApi.Helper module to be imported.
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

        [System.Collections.Generic.List[string]]$FileProjectFailed = New-Object 'System.Collections.Generic.List[string]'
        [System.Collections.Generic.List[string]]$FileProjectPassed = New-Object 'System.Collections.Generic.List[string]'
        
        foreach ($jsonFile in $Files) {
            Write-Verbose -Message "The current file is [$jsonFile]" -Verbose
            try {
                [PSCustomObject]$jsonObject = Get-JsonContent -Filename $jsonFile -PropertyName "variables"
                [psobject]$jsonContent = $jsonObject.JsonContent

                if ($jsonObject.ValidJson) {
                    
                    # Read and validate each JSON file and either Add, Update or ignore the Variablegroup
                    foreach ($project in $Projects) {
                       
                        if ($ProjectExclusions.Contains($project)) {
                            Write-Verbose -Message "Excluding the current project [$project] as requested" -Verbose
                            continue
                        }
                        Write-Verbose -Message "The current team project name is [$project]" -Verbose

                        [string]$apiGetUri = Get-VariablegroupGetUrl -Project $project -GroupName $jsonObject.Name
                        
                        try{
                            #Get all Variable Group
                            [psobject]$vstsVariableGroup = Invoke-RestMethod -Uri $apiGetUri -Method GET -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken); ContentType = ("application/json"); }
                        
                            if ($vstsVariableGroup.count -eq 0) {
                                #Create Variable Group
                                [string]$ApiPostUri = Get-VariablegroupPostUrl -Project $project
                                Invoke-RestMethod -Uri $ApiPostUri -Method POST -Body $jsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) } | Out-Null
                            }
                            else {
                                [string]$vgid = $vstsVariableGroup.value[0].id
                                [psobject]$variableGroupObject = $jsonObject.JsonObject
        
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
                            }

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
        [void]$SummaryBuilder.Append("    Variablegroup distribution summary report   `n")
        [void]$SummaryBuilder.Append("------------------------------------------------`n")
        [void]$SummaryBuilder.Append("`n")

        # Projects found in organization
        $organizationName = Get-OrganizationName
        if (-not [string]::IsNullOrEmpty($Projects) -and $Projects.Count -gt 0) {
            [void]$SummaryBuilder.Append([string]::Format("`nRun completed in [{0}] projects: [{1}]", $Projects.Count, $([System.String]::Join(', ', ($Projects | Sort-Object)))))
        }
        else {
            [void]$SummaryBuilder.Append("`nProject(s) not found in organization '$organizationName'")
        }
        [void]$SummaryBuilder.Append("`n")

        # Input files
        [void]$SummaryBuilder.Append([string]::Format("`n[{0}] JSON files (Variablegroup) were processed. They are`n", $Files.Count))
        ForEach ($fileName in $Files){ 
            [void]$SummaryBuilder.Append([string]::Format("{0}",[System.IO.Path]::GetFileName($fileName)))
            [void]$SummaryBuilder.Append("`n")
        }

        # Variablegroup distribution execution summary dashboard: Invalid input files        
        if (-not [string]::IsNullOrEmpty($FilesSkipped) -and $FilesSkipped.Count -gt 0) {
            $SkippedFileCount = $FilesSkipped.Count
            [void]$SummaryBuilder.Append("`nFollowing input files($SkippedFileCount) were ignored due to invalid Variablegroup definition.`n")
            foreach ($fileskipped in $FilesSkipped) {
                [void]$SummaryBuilder.Append([string]::Format("{0}",[System.IO.Path]::GetFileName($fileskipped)))
                [void]$SummaryBuilder.Append("`n")
            }
        }

        # Variablegroup distribution execution summary dashboard: successful projects
        if (-not [string]::IsNullOrEmpty($SuccededProjectCollection) -and $SuccededProjectCollection.Count -gt 0) {
            [void]$SummaryBuilder.Append("`nVariablegroup create/update successfully for input file: [projects] `n")
            ForEach ($itemKey in $SuccededProjectCollection.Keys){ 
                $itemValue = '[' + $SuccededProjectCollection[$itemKey] + ']'
                [void]$SummaryBuilder.Append([string]::Format("{0} : {1}",[System.IO.Path]::GetFileName($itemKey), $itemValue))
                [void]$SummaryBuilder.Append("`n")
            }
            [void]$SummaryBuilder.Append("`n")
        }

        # Variablegroup distribution execution summary dashboard: failed projects
        if (-not [string]::IsNullOrEmpty($FailedProjectCollection) -and $FailedProjectCollection.Count -gt 0) {
            [void]$SummaryBuilder.Append("`nVariablegroup create/update failed for input file: [projects]`n")
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
            [void]$SummaryBuilder.Append("`nOne or more projects failed while distributing Variablegroups. Details of errors can be found below `n")
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
    Write-Verbose "Maintain Pipelines distribute Variablegroup execution ended" -Verbose
    if (-not [string]::IsNullOrEmpty($FailedProjectCollection) -and $FailedProjectCollection.Count -gt 0){
        throw "One or more error(s) encountered during the execution. Please refer the summary above."
    }
}

Set-StrictMode -Off
Exit $Exitcode