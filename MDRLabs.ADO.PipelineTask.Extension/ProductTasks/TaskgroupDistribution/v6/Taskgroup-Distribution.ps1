<#
     .Synopsis
        Add or update an Azure DevOps Task Group. 

     .Description
        The script calls the taskgroups REST API to create a new, or update an existing Azure DevOps
        task group. The scipt is packaged in as an Azure DevOps Build Pipeline extension.

    .Parameter Token
    .Parameter Pat
        The user's Personal Access Token (PAT) with the correct scopes for creating and updating Taskgroups.

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
        One or more JSON files containing a valid taskgroup definition. 
        If the files are not in the same directory as the script they must be fully qualified.
        If providing multiple files, use a comma (,) to separate each.

    .Parameter RequestSleepTime
    .Parameter Sleep
        This is the time in seconds to pause while retrieving the Access Control List (ACL) for a specific
        Security Namespace, Project and Taskgroup. Potential workaround because retrieving permissions of a
        newly added task group can fail on the initial call to retrieve them

    .Parameter DenyContributorEditPermission
        Taskgroup permissions are used to grant edit, delete and administrator access to various Azure DevOps 
        groups like Contributers, Project Administrators, Release Administrators, Build Administrators

    .Inputs
        See Jsonfiles and Folder parameters

    .Outputs
        Exitcode 0 if successful, otherwise 1

    .Example
        Run the script from a PowerShell CLI
          Maintain-Pipelines [-verbose] [-RequestSleepTime seconds] -Token pat [-Projects "teamproject1","teamproject2" ... | -Exclude "teamproject1","teamproject2" ...] -Folder directoryname | -Files "file1","file2","file3" ...
        Run the script an from Azure DevOps PowerShell task
          Maintain-Pipelines [-debug] [-verbose] [-RequestSleepTime seconds] -Token pat [-Projects "teamproject1","teamproject2" ... | -Exclude "teamproject1","teamproject2" ...] -Folder directoryname | -Files "file1","file2","file3" ...

    .Link
        https://docs.microsoft.com/en-us/rest/api/azure/devops/?view=azure-devops-rest-5.1
        https://docs.microsoft.com/en-us/azure/devops/integrate/concepts/integration-bestpractices?view=vsts
        https://docs.microsoft.com/en-us/rest/api/vsts/distributedtask/taskgroups
        https://github.com/Microsoft/azure-pipelines-task-lib/blob/master/tasks.schema.json
        https://martin77s.wordpress.com/2014/06/17/powershell-scripting-best-practices/
        https://docs.microsoft.com/en-us/azure/devops/integrate/concepts/rate-limits?view=azure-devops&amp%3Btabs=new-nav&viewFallbackFrom=vsts&tabs=new-nav
        https://docs.microsoft.com/en-us/azure/devops/pipelines/library/task-groups?view=azure-devops
        https://docs.microsoft.com/en-us/azure/devops/organizations/security/permissions?view=azure-devops&tabs=preview-page

.Notes
        1. If running the script as an Azure DevOps extension the -Token variable should be stored as a secret variable.
        2. The -Projects and the -Exclude parameters are mutually exclusive. If the run contains a list of Team 
           Project names it doesn't make sense to also specify an Exclude list. Only include the latter if no
           -Project parameter is specified and there is a need to exclude ome or more team projects from the run.
        3. The -Folder and -Files parameters are mutually exclusive. If specified the -Folder parameter must 
           resolve to a valid directory containing JSON files, otherwise the the files specified with -Files
           will be used.
        4. The target project(s) is/are expected to have endpoints defined for SonarQube-Production. A regex mask is
           used to identify the correct endpoint. This is the regex: (?i)sonarqube(?:\s+|[-_])(?:pr)(?:oduction){0,1}
        5. The target project(s) is/are expected to have endpoints defined for Splunk-Production. A regex mask is
           used to identify the correct endpoint. This is the regex: (?i)splunk(?:\s+|[-_])(?:pr)(?:oduction){0,1}
        6. RATE LIMITS
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

    # RequestSleepTime
    [Parameter(Mandatory = $false, HelpMessage = 'This is default time in seconds (8) to pause distribution while retrieving the Access Control List (ACL) for a specific
    Security Namespace, Project and Taskgroup.')]
    [Alias('Sleep')]
    [string]
    $RequestSleepTime = "5",

    # Projects / AzureDevops Team Projects parameter
    [Parameter(Mandatory = $true, ParameterSetName = 'Projects', HelpMessage = 'A comma/quoted list of one or more team project names')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Folder')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Files')]
    [AllowEmptyCollection()]
    [Alias('Projects')]
    [string[]]
    $AzureDevOpsTeamProjects,

    # Folder parameter
    [Parameter(Mandatory = $true, ParameterSetName = 'Folder', HelpMessage = 'The name of a folder containing Taskgroups')]
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
    [Parameter(Mandatory = $true, ParameterSetName = 'Files', HelpMessage = 'A comma/quoted list of JSON Taskgroup files')]
    [AllowEmptyCollection()]
    [Alias('Files')]
    [string[]]
    $Jsonfiles,

    # DenyContributorEditPermission
    [Parameter(Mandatory = $false, HelpMessage = 'Deny the Contributor role Edit permission on distributed task groups')]
    [bool]
    $DenyContributorEditPermission = 0
)

#******************************************************************************************************************
# Script body
# Execution begins here
#******************************************************************************************************************
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop';

try {
    $global:ScriptFolder = Split-Path -Parent $myInvocation.MyCommand.Path

    Write-Verbose -Message ("Taskgroup-Distribution run has started") -Verbose
    Write-Verbose -Message ("The PowerShell version is [{0}]" -f $PSVersionTable.PSVersion)
    Write-Verbose -Message ("The chosen parameter set name is [{0}]" -f $PSCmdlet.ParameterSetName)

    [int]$Exitcode = 0
    [psobject]$JsonContent = @{}
    [System.Collections.Generic.Dictionary[string, string]]$ProjectDictionary = New-Object 'System.Collections.Generic.Dictionary[string,string]'
    [System.Collections.Generic.List[string]]$Projects = New-Object 'System.Collections.Generic.List[string]'
    [System.Collections.Generic.List[string]]$ProjectExclusions = New-Object 'System.Collections.Generic.List[string]'
    [System.Collections.Generic.List[string]]$Files = New-Object 'System.Collections.Generic.List[string]'

    # Execution summary arrays
    [System.Collections.Generic.List[string]]$FilesSkipped = New-Object 'System.Collections.Generic.List[string]'
    $FailedProjectCollection = @{}
    $SuccededProjectCollection = @{}

    # Expects the DesAzure.PowerShell.Vsts.RestApi.Helper module to be imported.
    if (-not(Get-Module -name "MDRLabs.Ps.Ado.RestApi.HelperMDRLabs.ADORestApiHelper")) {
        $localmodules = Join-Path $ScriptFolder "Modules\MDRLabs.Ps.Ado.RestApi.HelperMDRLabs.ADORestApiHelper.psm1"
        import-module $localmodules -Force -NoClobber -Global -ErrorAction Stop
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
        Write-Warning -Message "The current run has no valid JSON files (taskgroups) to distribute. Terminating the run" -Verbose
    }
    else {  
        Write-Verbose -Message ("There are [{0}] JSON files (taskgroups) to process. They are [{1}]" -f $Files.Count, $([System.String]::Join(', ', ($Files | Sort-Object))))

        [System.Collections.Generic.List[string]]$FileProjectFailed = New-Object 'System.Collections.Generic.List[string]'
        [System.Collections.Generic.List[string]]$FileProjectPassed = New-Object 'System.Collections.Generic.List[string]'
        [string] $ErrorMessageBuilder = ""
        foreach ($jsonFile in $Files) {

            Write-Verbose -Message "The current file is [$jsonFile]" -Verbose
            try {

                [PSCustomObject]$jsonObject = Get-JsonContent -Filename $jsonFile -DefinitionType "metaTask"
                if ($jsonObject.ValidJson) {

                    [psobject]$jsonContent = $jsonObject.JsonContent
                    [psobject]$jsonTaskgroupObject = $jsonObject.JsonObject
                    [string]$jsonObjectName = $jsonObject.Name

                    # Identify the security-namespace for MetaTask
                    foreach ($project in $Projects) {
                        
                        if ($ProjectExclusions.Contains($project)) {
                            Write-Verbose -Message "Excluding the current project [$project] as requested" -Verbose
                            continue
                        }

                        [string]$apiTaskgroupPostUrl = Get-TaskgroupPostUrl -Project $project
                        [string]$apiTaskgroupGetUrl = Get-TaskgroupGetUrl -Project $project
                        
                        Write-Verbose -Message "The current team project name is [$project]" -Verbose
                        Write-Verbose -Message ("Taskgroup Post Url is [$apiTaskgroupPostUrl]")
                        Write-Verbose -Message ("Taskgroup Get Url is [$apiTaskgroupGetUrl]")

                        [psobject]$taskGroupList = Invoke-RestMethod -Uri $apiTaskgroupGetUrl -Method GET -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken); ContentType = ("application/json"); }
                        [psobject]$uniqueListOfTaskgroups = ($taskGroupList.value | Group-Object 'name' | ForEach-Object { $_.Group | select-object 'name', 'revision', 'version', 'id' -First 1 | sort-object 'name'})

                        # Check if the current taskgroup already exists in the current project.
                        if ($taskGroupList.count -eq 0 -or -not ($uniqueListOfTaskgroups | Where-Object -FilterScript {$_.name -eq $jsonObjectName})) {
                            [psobject]$updatedJsonContent = $jsonTaskgroupObject | ConvertTo-Json -Depth 10 -Compress
                            Write-Verbose -Message "Taskgroup doesn't exist. Creating [$jsonObjectName]" -Verbose
                        
                            try {
                                # Create project taskgroup
                                $null = Invoke-RestMethod -Uri $apiTaskgroupPostUrl -Method POST -Body $updatedJsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) }
                                Write-Verbose -Message "Task Group [$jsonObjectName] for Team project [$project] has been created" -Verbose
                            
                                if($true -eq $DenyContributorEditPermission)
                                {
                                    # Apply project permissions to taskgroup
                                    [PSCustomObject]$responseObject = Set-TaskGroupPermissions -Project $project -TaskgroupName $jsonObjectName -RequestSleepTime $RequestSleepTime
                                    Write-Verbose -Message ("The Set Permissions status for project [{0}], taskgroup [{1}] was [{2}]" -f $project, $jsonObjectName, $responseObject.SetPermissions) -Verbose
                                }
                                $FileProjectPassed.Add($project)
                            }
                            catch [Exception]{
                                $FileProjectFailed.Add($project)
                                if($null -ne $_.ErrorDetails){
                                    $ErrorMessageBuilder = $ErrorMessageBuilder + "JsonFile=$jsonFile, Project=$project, Error Message=" + $_.ErrorDetails.Message + "`r`n"
                                    Write-Error -Message ("Caught an exception in project [{0}] of file [{1}] of type [{2}], reason [{3}], details [{4}]" -f $project, $jsonFile, $_.Exception.GetType().FullName, $_.Exception.Message, $_.ErrorDetails.Message) -ErrorAction Continue
                                }
                                else {
                                    $ErrorMessageBuilder = $ErrorMessageBuilder + "JsonFile=$jsonFile, Project=$project, Error Message=" + $_.Exception.Message + "`r`n"
                                    Write-Error -Message ("Caught an exception in project [{0}] of file [{1}] of type [{2}], reason [{3}]" -f $project, $jsonFile, $_.Exception.GetType().FullName, $_.Exception.Message) -ErrorAction Continue
                                }
                             }
                        }

                        # Update taskgroups
                        foreach ($remoteTaskgroup in $uniqueListOfTaskgroups){
                            if ($jsonObjectName -eq $remoteTaskgroup.name) {                                 
                                try {
                                        #region Update taskgroup
                                        if($jsonTaskgroupObject.PSObject.Properties['isMinorChange'] -and $jsonTaskgroupObject.isMinorChange -eq $false) {
                                            if ($jsonTaskgroupObject.PSObject.Properties.Match("parentDefinitionId").Count) {
                                                $jsonTaskgroupObject.parentDefinitionId = $remoteTaskgroup.id
                                            }
                                            else {
                                                $jsonTaskgroupObject | Add-Member -MemberType NoteProperty -Name 'parentDefinitionId' -Value $remoteTaskgroup.id
                                            }
            
                                            #region create draft
                                            [psobject]$updatedJsonContent = $jsonTaskgroupObject | ConvertTo-Json -Depth 10 -Compress
                                            [psobject]$responseDraft = Invoke-RestMethod -Uri $apiTaskgroupPostUrl -Method POST -Body $updatedJsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) }
                                            #endregion create draft
            
                                            #region publish draft, create preview
                                            [string]$apiTaskgroupPutUrl = Get-TaskgroupPutUrl -Project $project -ParentTaskgroupId $remoteTaskgroup.id
                                            [pscustomobject]$changeSet = @{
                                                "parentDefinitionRevision" = $remoteTaskgroup.revision
                                                "preview"                  = "true"
                                                "taskGroupId"              = $responseDraft.id
                                                "taskGroupRevision"        = $responseDraft.revision
                                            }
                                            [psobject]$changeReference = $changeSet | ConvertTo-Json -Depth 10 -Compress
                                            [psobject]$responsePreview = Invoke-RestMethod -Uri $apiTaskgroupPutUrl -Method PUT -Body $changeReference -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) }
                                            #endregion publish draft, create preview
            
                                            #region publish preview
                                            [string]$apiTaskgroupPatchUrl = Get-TaskgroupGetUrl -Project $project -TaskgroupId $remoteTaskgroup.id -DisablePriorVersions 'false'
                                            $changeSet = @{
                                                "comment"  = "Publish Preview: $CommitMessage"
                                                "preview"  = "false"
                                                "version"  = $responsePreview.value[0].version
                                                "revision" = $responsePreview.value[0].revision
                                            }
                                            $changeReference = $changeSet | ConvertTo-Json -Depth 10 -Compress
                                            $null = Invoke-RestMethod -Uri $apiTaskgroupPatchUrl -Method PATCH -Body $changeReference -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) }
                                            #endregion publish preview     
            
                                            Write-Verbose -Message ("Taskgroup [{0}] has been updated for Team project [{1}] has been updated" -f $jsonObjectName, $project) -Verbose
                                            Write-Verbose -Message ("Taskgroup [{0}] old version [{1}], new version [{2}] for Team project [{3}] has been updated" -f $jsonObjectName, $remoteTaskgroup.version, $responsePreview.value[0].version, $project)
                                        }
                                        else {
                                            [string]$uuid = $remoteTaskgroup.id  
                                            [int]$revision = $remoteTaskgroup.revision 
                                            [psobject]$version = $remoteTaskgroup.version                             
                                            if ( $null -ne $jsonTaskgroupObject.id ) {
                                                $jsonTaskgroupObject.id = $uuid
                                            }
                                            else {
                                                $jsonTaskgroupObject | Add-Member -MemberType NoteProperty -Name 'id' -Value $uuid
                                            }
                                            if ( $null -ne $jsonTaskgroupObject.revision ) {
                                                $jsonTaskgroupObject.revision = $revision
                                            }
                                            else {
                                                $jsonTaskgroupObject | Add-Member -MemberType NoteProperty -Name 'revision' -Value $revision
                                            }
                                            if ( $null -ne $jsonTaskgroupObject.version ) {
                                                $jsonTaskgroupObject.version = $version
                                            }
                                            else {
                                                $jsonTaskgroupObject | Add-Member -MemberType NoteProperty -Name 'version' -Value $version
                                            }
                                            [psobject]$updatedJsonContent = $jsonTaskgroupObject | ConvertTo-Json -Depth 10 -Compress
                                            [string]$apiTaskgroupPutUrl = Get-TaskgroupPutUrl -Project $project -TaskgroupId $remoteTaskgroup.id
                                            [psobject]$responsePreview = Invoke-RestMethod -Uri $apiTaskgroupPutUrl -Method PUT -Body $updatedJsonContent -ContentType 'application/json' -Headers @{ Authorization = ("{0} {1}" -f $AuthorisationType, $EncodedToken) }
            
                                            Write-Verbose -Message ("Taskgroup [{0}] has been updated for Team project [{1}] has been updated" -f $jsonObjectName, $project) -Verbose
                                            Write-Verbose -Message ("Taskgroup [{0}] old version [{1}], new version [{2}] for Team project [{3}] has been updated" -f $jsonObjectName, $remoteTaskgroup.version, $responsePreview.version, $project)
                                        }
                                        #endregion Update taskgroup

                                        if($true -eq $DenyContributorEditPermission)
                                        {
                                            #region Apply project permissions to taskgroup                            
                                            [PSCustomObject]$responseObject = Set-TaskGroupPermissions -Project $project -TaskgroupName $jsonObjectName -RequestSleepTime $RequestSleepTime
                                            Write-Verbose -Message ("The Set Permissions status for project [{0}], taskgroup [{1}] was [{2}]" -f $project, $jsonObjectName, $responseObject.SetPermissions) -Verbose
                                            #endregion Apply project permissions to taskgroup
                                        }
                                 
                                        $FileProjectPassed.Add($project)
                                }
                                catch [Exception]{
                                    $FileProjectFailed.Add($project)
                                    if($null -ne $_.ErrorDetails){
                                        $ErrorMessageBuilder = $ErrorMessageBuilder + "JsonFile=$jsonFile, Project=$project, Error Message=" + $_.ErrorDetails.Message + "`r`n"
                                        Write-Error -Message ("Caught an exception in project [{0}] of file [{1}] of type [{2}], reason [{3}], details [{4}]" -f $project, $jsonFile, $_.Exception.GetType().FullName, $_.Exception.Message, $_.ErrorDetails.Message) -ErrorAction Continue
                                    }
                                    else {
                                        $ErrorMessageBuilder = $ErrorMessageBuilder + "JsonFile=$jsonFile, Project=$project, Error Message=" + $_.Exception.Message + "`r`n"
                                        Write-Error -Message ("Caught an exception in project [{0}] of file [{1}] of type [{2}], reason [{3}]" -f $project, $jsonFile, $_.Exception.GetType().FullName, $_.Exception.Message) -ErrorAction Continue
                                    }
                                }
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
                    $ErrorMessageBuilder = $ErrorMessageBuilder + "JsonFile=$jsonFile, Error Message=" + $_.ErrorDetails.Message + "`r`n"
                    Write-Error -Message ("Caught an exception while reading taskgroup '$jsonFile' of type [{0}], reason [{1}], details [{2}]" -f $_.Exception.GetType().FullName, $_.Exception.Message, $_.ErrorDetails.Message) -ErrorAction Continue
                }
                else {
                    $ErrorMessageBuilder = $ErrorMessageBuilder + "JsonFile=$jsonFile, Error Message=" + $_.Exception.Message + "`r`n"
                    Write-Error -Message ("Caught an exception while reading taskgroup '$jsonFile' of type [{0}], reason [{1}]" -f $_.Exception.GetType().FullName, $_.Exception.Message) -ErrorAction Continue
                }
            }
        }

        Write-Verbose -Message "------------------------------------------------"
        Write-Verbose -Message "    Taskgroup distribution summary report   "
        Write-Verbose -Message "------------------------------------------------"
        Write-Verbose -Message ""

        # Projects found in organization
        $organizationName = Get-OrganizationName
        if (-not [string]::IsNullOrEmpty($Projects) -and $Projects.Count -gt 0) {
            Write-Verbose -Message ("Run completed in [{0}] projects: [{1}]" -f $Projects.Count, $([System.String]::Join(', ', ($Projects | Sort-Object)))) -Verbose
        }
        else {
            Write-Verbose -Message "Project(s) not found in organization '$organizationName'"
        }
        Write-Verbose -Message ""

        # Input files
        Write-Verbose -Message ("[{0}] JSON files (taskgroups) were processed. They are [{1}]" -f $Files.Count, $([System.String]::Join(', ', ($Files | Sort-Object))))
        Write-Verbose -Message ""

        # Taskgroup distribution execution summary dashboard: Invalid input files        
        if (-not [string]::IsNullOrEmpty($FilesSkipped) -and $FilesSkipped.Count -gt 0) {
            $SkippedFileCount = $FilesSkipped.Count
            Write-Verbose -Message "Following input files($SkippedFileCount) were ignored due to invalid taskgroup definition."
            foreach ($fileskipped in $FilesSkipped) {
                Write-Verbose -Message "[$fileskipped]"
            }
            Write-Verbose -Message ""
        }

        # Taskgroup distribution execution summary dashboard: successful projects
        if (-not [string]::IsNullOrEmpty($SuccededProjectCollection) -and $SuccededProjectCollection.Count -gt 0) {
            Write-Verbose -Message "Taskgroups create/update successfully for input file: [projects]"
            ForEach ($itemKey in $SuccededProjectCollection.Keys){ 
                $itemValue = '[' + $SuccededProjectCollection[$itemKey] + ']'
                Write-Verbose -Message "$itemKey : $itemValue"
                Write-Verbose -Message ""
            }
            Write-Verbose -Message ""
        }

        # Taskgroup distribution execution summary dashboard: failed projects
        if (-not [string]::IsNullOrEmpty($FailedProjectCollection) -and $FailedProjectCollection.Count -gt 0) {
            Write-Verbose -Message "Taskgroups create/update failed for input file: [projects]"
            ForEach ($itemKey in $FailedProjectCollection.Keys){ 
                $itemValue = '[' + $FailedProjectCollection[$itemKey] + ']'
                Write-Verbose -Message "$itemKey : $itemValue"
                Write-Verbose -Message ""
            }
            Write-Verbose -Message ""
        }

        # Display all error message text
        if(-not [string]::IsNullOrEmpty($ErrorMessageBuilder))
        {
            # Mark pipeline as failure and dissplay all errors 
            Write-Verbose -Message ""
            Write-Verbose -Message "One or more projects failed while distributing taskgroups. Details of errors can be found below" 
            Write-Verbose -Message $ErrorMessageBuilder 
            $Exitcode = 1
        }

        Write-Verbose -Message "------------------------------------------------"
    }
}
catch [Exception] {
    if($null -ne $_.ErrorDetails){
        Write-Error -Message ("Caught an exception of type [{0}], reason [{1}], details [{2}]" -f $_.Exception.GetType().FullName, $_.Exception.Message, $_.ErrorDetails.Message) -ErrorID "1" -Targetobject $_
    }
    else {
        Write-Error -Message ("Caught an exception of type [{0}], reason [{1}]" -f $_.Exception.GetType().FullName, $_.Exception.Message) -ErrorID "1" -Targetobject $_
    }
    $Exitcode = 1
}
finally {
    Write-Verbose "Distribution-Pipelines task execution ended" -Verbose
    if (-not [string]::IsNullOrEmpty($FailedProjectCollection) -and $FailedProjectCollection.Count -gt 0){
        throw "One or more error(s) encountered during the execution. Please refer the summary above."
    }
}

Set-StrictMode -Off
Exit $Exitcode
