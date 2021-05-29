#
#.SYNOPSIS
#	Runs a PowerShell command with output to VSTS

#.DESCRIPTION
#	Runs a PowerShell command with output to VSTS

#.OUTPUTS
#	Progress messages
#*/

$rollForwardTable = @{
    "5.0.0" = "5.1.1";
};

Function Get-SavedModulePath {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]        
        [String] $AzurePowerShellVersion,
        [Switch] $Classic
    )
    
    If ($Classic -eq $true) {
        Return $($env:SystemDrive + "\Modules\Azure_" + $AzurePowerShellVersion)
    }

    Return $($env:SystemDrive + "\Modules\AzureRm_" + $AzurePowerShellVersion)
}

Function Update-PSModulePathForHostedAgent {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]   
        [String] $TargetAzurePs,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]           
        [String] $AuthScheme
    )
    Trace-VstsEnteringInvocation $MyInvocation
    Try {
        If ($TargetAzurePs) {
            $hostedAgentAzureRmModulePath = Get-SavedModulePath -azurePowerShellVersion $TargetAzurePs
            $hostedAgentAzureModulePath = Get-SavedModulePath -azurePowerShellVersion $TargetAzurePs -Classic
        }
        Else {
            $hostedAgentAzureRmModulePath = Get-LatestModule -PatternToMatch "^azurerm_[0-9]+\.[0-9]+\.[0-9]+$" -PatternToExtract "[0-9]+\.[0-9]+\.[0-9]+$" -Classic:$false
            $hostedAgentAzureModulePath  =  Get-LatestModule -PatternToMatch "^azure_[0-9]+\.[0-9]+\.[0-9]+$"   -PatternToExtract "[0-9]+\.[0-9]+\.[0-9]+$" -Classic:$true
        }

        If ($AuthScheme -eq 'ServicePrincipal' -or $AuthScheme -eq 'ManagedServiceIdentity' -or $AuthScheme -eq '') {
            $env:PSModulePath = $hostedAgentAzureModulePath + ";" + $env:PSModulePath
            $env:PSModulePath = $env:PSModulePath.TrimStart(';')
            $env:PSModulePath = $hostedAgentAzureRmModulePath + ";" + $env:PSModulePath
            $env:PSModulePath = $env:PSModulePath.TrimStart(';')
        }
        Else {
            $env:PSModulePath = $hostedAgentAzureRmModulePath + ";" + $env:PSModulePath
            $env:PSModulePath = $env:PSModulePath.TrimStart(';')
            $env:PSModulePath = $hostedAgentAzureModulePath + ";" + $env:PSModulePath
            $env:PSModulePath = $env:PSModulePath.TrimStart(';')
        }
    } Finally {
        Write-Verbose "The updated value of the PSModulePath is: $($env:PSModulePath)"
        Trace-VstsLeavingInvocation $MyInvocation
    }
}

Function Get-LatestModule {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]           
        [String] $PatternToMatch,
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]           
        [String] $PatternToExtract,
        [Switch] $Classic
    )
    
    $resultFolder = ""
    $regexToMatch = New-Object -TypeName System.Text.RegularExpressions.Regex -ArgumentList $PatternToMatch
    $regexToExtract = New-Object -TypeName System.Text.RegularExpressions.Regex -ArgumentList $PatternToExtract
    $maxVersion = [version] "0.0.0"

    Try {
        $moduleFolders = Get-ChildItem -Directory -Path $($env:SystemDrive + "\Modules") | Where-Object { $regexToMatch.IsMatch($_.Name) }
        ForEach ($moduleFolder in $moduleFolders) {
            $moduleVersion = [version] $($regexToExtract.Match($moduleFolder.Name).Groups[0].Value)
            If ($moduleVersion -gt $maxVersion) {
                If ($Classic) {
                    $modulePath = [System.IO.Path]::Combine($moduleFolder.FullName,"Azure\$moduleVersion\Azure.psm1")
                } 
                Else {
                    $modulePath = [System.IO.Path]::Combine($moduleFolder.FullName,"AzureRM\$moduleVersion\AzureRM.psm1")
                }

                If (Test-Path -LiteralPath $modulePath -PathType Leaf) {
                    $maxVersion = $moduleVersion
                    $resultFolder = $moduleFolder.FullName
                } 
                Else {
                    Write-Verbose "A folder matching the module folder pattern was found at $($moduleFolder.FullName) but didn't contain a valid module file"
                }
            }
        }
    }
    Catch {
        Write-Verbose "Attempting to find the Latest Module Folder failed with the error: $($_.Exception.Message)"
        $resultFolder = ""
    }
    Write-Verbose "Latest module folder detected: $resultFolder"
    Return $resultFolder
}

Function Get-RollForwardVersion {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $AzurePowerShellVersion
    )

    Trace-VstsEnteringInvocation $MyInvocation
    
    Try {
        $rollForwardAzurePSVersion = $rollForwardTable[$AzurePowerShellVersion]
        If (![string]::IsNullOrEmpty($rollForwardAzurePSVersion)) {
            $hostedAgentAzureRmModulePath = Get-SavedModulePath -AzurePowerShellVersion $rollForwardAzurePSVersion
            $hostedAgentAzureModulePath = Get-SavedModulePath -AzurePowerShellVersion $rollForwardAzurePSVersion -Classic
        
            If ((Test-Path -Path $hostedAgentAzureRmModulePath) -eq $true -or (Test-Path -Path $hostedAgentAzureModulePath) -eq $true) {
                Write-Warning (Get-VstsLocString -Key "OverrideAzurePowerShellVersion" -ArgumentList $AzurePowerShellVersion, $rollForwardAzurePSVersion)
                Return $rollForwardAzurePSVersion;
            }
        }
        Return $AzurePowerShellVersion
    }
    Finally {
        Trace-VstsLeavingInvocation $MyInvocation
    }
}

Function Run-PsScript {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String] $ScriptCommand
    )

    # Remove all commands imported from VstsTaskSdk, other than Out-Default.
    # Remove all commands imported from VstsAzureHelpers_.
    Get-ChildItem -LiteralPath function: |
        Where-Object {
            ($_.ModuleName -eq 'VstsTaskSdk' -and $_.Name -ne 'Out-Default') -or
            ($_.Name -eq 'Invoke-VstsTaskScript') -or
            ($_.ModuleName -eq 'VstsAzureHelpers_' )
        } |
        Remove-Item

    # For compatibility with the legacy handler implementation, set the error action
    # preference to continue. An implication of changing the preference to Continue,
    # is that Invoke-VstsTaskScript will no longer handle setting the result to failed.
    $global:ErrorActionPreference = 'Continue'

    # Undocumented VstsTaskSdk variable so Verbose/Debug isn't converted to ##vso[task.debug].
    # Otherwise any content the ad-hoc script writes to the verbose pipeline gets dropped by
    # the agent when System.Debug is not set.
    $global:__vstsNoOverrideVerbose = $true

    # Run the user's script. Redirect the error pipeline to the output pipeline to enable
    # a couple goals due to compatibility with the legacy handler implementation:
    # 1) STDERR from external commands needs to be converted into error records. Piping
    #    the redirected error output to an intermediate command before it is piped to
    #    Out-Default will implicitly perform the conversion.
    # 2) The task result needs to be set to failed if an error record is encountered.
    #    As mentioned above, the requirement to handle this is an implication of changing
    #    the error action preference.
    ([scriptblock]::Create($ScriptCommand)) | 
        ForEach-Object {
            Remove-Variable -Name scriptCommand -ErrorAction SilentlyContinue
            Write-Host "##[command]$_"
            . $_ 2>&1
        } | 
        ForEach-Object {
            If ($_ -is [System.Management.Automation.ErrorRecord]) {
                If ($_.FullyQualifiedErrorId -eq "NativeCommandError" -or $_.FullyQualifiedErrorId -eq "NativeCommandErrorMessage") {
                    ,$_
                    If ($__vsts_input_failOnStandardError -eq $true) {
                        "##vso[task.complete result=Failed]"
                    }
                }
                Else {
                    If ($__vsts_input_errorActionPreference -eq "continue") {
                        ,$_
                        If ($__vsts_input_failOnStandardError -eq $true) {
                            "##vso[task.complete result=Failed]"
                        }
                    }
                    ElseIf ($__vsts_input_errorActionPreference -eq "stop") {
                        Throw $_
                    }
                }
            }
            Else {
                ,$_
            }
        }
}