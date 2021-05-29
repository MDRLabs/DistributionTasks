# Synopsis

Add or update an Azure DevOps Task Group.

## Description

The script calls the taskgroups REST API to create a new, or update an existing Azure DevOps ttask group. The scipt is packaged in as an Azure DevOps Build Pipeline extension.

## Parameter Token Alias Pat

The user's Personal Access Token (PAT) with the correct scopes for creating and updating Taskgroups.

## Parameter CommitMessage Alias Comment

The reason for the change. This would normally be the last commit message

## Parameter AzureDevOpsTeamProjects Alias Projects

Projects are the Azure DevOps Team Project(s) to target. If no value is specified, all Azure DevOps Team Projects will be targeted. If providing multiple teams, use a comma (,) to separate each and quote the names if they contain spaces.

## Parameter ExcludeTeamProjects Alias Exclude

To exclude one or more Azure DevOps Team Projects from the run use this parameter. If providing multiple teams, use a comma (,) to separate each and quote the names if they contain spaces.

## Parameter Folder

This is the name of a single directory. Unless the value contains a fully qualified directory path, the directory is assumed to be relative to the script folder. Only JSON files in this folder will be included in the run.

## Parameter Jsonfiles Alias Files

One or more JSON files containing a valid taskgroup definition. If the files are not in the same directory as the script
they must be fully qualified. If providing multiple files, use a comma (,) to separate each.

## Parameter RequestSleepTime Alias Sleep

This is the time in seconds to pause while retrieving the Access Control List (ACL) for a specific Security Namespace, Project and Taskgroup. Potential workaround because retrieving permissions of a
newly added task group can fail on the initial call to retrieve them

## Parameter DenyContributorEditPermission

Taskgroup permissions are used to grant edit, delete and administrator access to various Azure DevOps 
groups like Contributers, Project Administrators, Release Administrators, Build Administrators. Check the checkbox if taskgroup permissions are required otherwise uncheck

## Inputs

See Jsonfiles and Folder parameters

## Outputs

Exitcode 0 if successful, otherwise 1

## Example

### Run the script from the Pipeline Task

Taskgroup-Distribution [-verbose] -Token pat [-RequestSleepTime seconds] [-Projects "teamproject1","teamproject2" ... | -Exclude "teamproject1","teamproject2" ...] -Folder directoryname | -Files "file1","file2","file3" ... 
| -DenyContributorEditPermission 0

### Run the script from a PowerShell CLI

Taskgroup-Distribution [-debug] [-verbose] -Token pat [-RequestSleepTime seconds] [-Projects "teamproject1","teamproject2" ... | -Exclude "teamproject1","teamproject2" ...] -Folder directoryname | -Files "file1","file2","file3" ... | -DenyContributorEditPermission 0

## Link

* [Azure DevOps REST API](https://docs.microsoft.com/en-us/rest/api/azure/devops/?view=azure-devops-rest-5.1)
* [Azure DevOps Integration Best Practices](https://docs.microsoft.com/en-us/azure/devops/integrate/concepts/integration-bestpractices?view=vsts)
* [Azure DevOps Taskgroups](https://docs.microsoft.com/en-us/rest/api/vsts/distributedtask/taskgroups)
* [PowerShell Credential Manager](https://github.com/davotronic5000/PowerShell_Credential_Manager)
* [PowerShell Best Practices](https://martin77s.wordpress.com/2014/06/17/powershell-scripting-best-practices/)
* [Rate Limits](https://docs.microsoft.com/en-us/azure/devops/integrate/concepts/rate-limits?view=azure-devops&amp%3Btabs=new-nav&viewFallbackFrom=vsts&tabs=new-nav)

## Notes

1. If running the script as an Azure DevOps extension the -Token variable should be stored as a secret variable.
2. The -Projects and the -Exclude parameters are mutually exclusive. If the run contains a list of Team 
   Project names it doesn't make sense to also specify an Exclude list. Only include the latter if no
   -Project parameter is specified and there is a need to exclude ome or more team projects from the
   run.
3. The -Folder and -Files parameters are mutually exclusive. If specified the -Folder parameter must 
   resolve to a valid directory containing JSON files, otherwise the the files specified with -Files
   will be used.
4. The DenyContributorEditPermission used to  to grant edit, delete and administrator access to various Azure DevOps 
   groups like Contributers, Project Administrators, Release Administrators, Build Administrators. 
5. The target project(s) is/are expected to have endpoints defined for SonarQube-Production. A regex mask is
   used to identify the correct endpoint. This is the regex: (?i)sonarqube(?:\s+|[-_])(?:pr)(?:oduction){0,1}