Trace-VstsEnteringInvocation $MyInvocation

Write-Verbose -Message ("Start Maintain Pipelines task execution")

# Dot source the private functions.
. $PSScriptRoot/Common/Run-PsScript.ps1

# Mandatory parameters
[string]$Token = Get-VstsInput -Name Token -Require
[string]$FileOrFolderSelector = Get-VstsInput -Verbose -Name FileOrFolderSelector -Require
[string]$TeamProjectsToTarget = Get-VstsInput -Verbose -Name TeamProjectsToTarget -Require
[string]$CommitMessage = Get-VstsInput -Verbose -Name CommitMessage -Require
[string]$RequestSleepTime = Get-VstsInput -Name RequestSleepTime -Require
[string]$ExtraVerbose = Get-VstsInput -Name ExtraVerbose -Require
[string]$DenyContributorEditPermission = Get-VstsInput -Name DenyContributorEditPermission -Require

# Optional parameters definition / initialisation
[string]$Comment = $CommitMessage.Replace('"', "").Replace("'", "")
[string]$Folder = "''"
[string[]]$Files = @()
[string[]]$Projects = @()
[string[]]$ExcludeTeamProjects = @()
[string]$FilesParam = @()
[string]$ProjectsParam = @()
[string]$ExcludeParam = @()
[string]$Parameters = "-Token $Token -Comment '$Comment' -RequestSleepTime $RequestSleepTime "

# If Verbosity is set to true show more messages
if ($ExtraVerbose -eq "True") {
    $Parameters += "-Verbose "
}

# If DenyContributorEditPermission is set to true then  apply project permissions to task group
if ($DenyContributorEditPermission -eq "True") {
    $Parameters += "-DenyContributorEditPermission 1 "
}
else {
    $Parameters += "-DenyContributorEditPermission 0 "
}

# What is the Source Type?
# The delimter is a comma (,). Remove quotes and extra spaces before or after a comma.
Switch ($FileOrFolderSelector) {
    "Folder" {
        $Folder = (Get-VstsInput -Verbose -Name Folder -Require)
        $Parameters += "-Folder $Folder "
    }
    "Files" {
        $Files = $Files + ( `
            (Get-VstsInput -Verbose -Name Files -Require).Replace('"', "").Replace("'", "") `
                -split "\s*,\s*" | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } `
        )
    }
    Default {
        Write-Error -Message ("Unknown FileOrFolderSelector option [$FileOrFolderSelector]")
    }
}

# What is the Source Type?
# The delimter is a comma (,). Remove quotes and extra spaces before or after a comma.
Switch ($TeamProjectsToTarget) {
    "TeamProjects" {
        $Projects = $Projects + ( `
            (Get-VstsInput -Verbose -Name Projects -Require).Replace('"', "").Replace("'", "") `
                -split "\s*,\s*" | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } `
        )
    }
    "Exclude" {
        $ExcludeTeamProjects = $ExcludeTeamProjects + ( `
            (Get-VstsInput -Verbose -Name ExcludeTeamProjects -Require).Replace('"', "").Replace("'", "") `
                -split "\s*,\s*" | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } `
        )
    }
    "Neither" {
        Write-Verbose -Message ("All Team Projects will be targetted")
    }
    Default {
        Write-Error -Message ("Unknown TeamProjectsToTarget option [$TeamProjectsToTarget]")
    }
}

# Prepare Collections to pass as input parameters to the script
# Quote and comma separate parameters
If ($Files.Count -gt 0) {
    $FilesParam = "'" + ($Files -join "', '") + "'"
    $Parameters += "-Files $FilesParam "
}
If ($Projects.Count -gt 0) {
    $ProjectsParam = "'" + ($Projects -join "', '") + "'"
    $Parameters += "-Projects $ProjectsParam "
}
If ($ExcludeTeamProjects.Count -gt 0) {
    $ExcludeParam = "'" + ($ExcludeTeamProjects -join "', '") + "'"
    $Parameters += "-Exclude $ExcludeParam "
}

Write-Verbose -Message ("Verbose: .................. [$ExtraVerbose]")
Write-Verbose -Message ("FileOrFolderSelector: ..... [$FileOrFolderSelector]")
Write-Verbose -Message ("TeamProjectsToTarget: ..... [$TeamProjectsToTarget]")
Write-Verbose -Message ("CommitMessage: ............ [$CommitMessage]")
Write-Verbose -Message ("Comment: .................. [$Comment]")
Write-Verbose -Message ("RequestSleepTime: ......... [$RequestSleepTime]")
Write-Verbose -Message ("Folder: ................... [$Folder]")
Write-Verbose -Message ("Files: .................... [$FilesParam]")
Write-Verbose -Message ("Projects: ................. [$ProjectsParam]")
Write-Verbose -Message ("ExcludeTeamProjects: ...... [$ExcludeParam]")
Write-Verbose -Message ("DenyContributorEditPermission: ...... [$DenyContributorEditPermission]")
Write-Verbose -Message ("Parameters: ............... [$Parameters]")

$scriptFile = Join-Path -Path $PSScriptRoot -ChildPath "Taskgroup-Distribution.ps1"
$scriptCommand = "& '$scriptFile' $Parameters "

Run-PsScript -ScriptCommand $scriptCommand 
