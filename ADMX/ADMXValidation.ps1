####################################################################
#                                                                  #
#          Script  : ADMXValidation                                #
#                                                                  #
#          Version : 1.1                                           #
#                                                                  #
#          Author  : Camille Debay                                 #
#                                                                  #
####################################################################

# TO DO 
# Export ADMX File without the problematic node


###############################################
# Script parameters
###############################################
param(
        [Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage="Path to ADMX files or folder containing ADMX files")][String] $Path,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true,HelpMessage="Display correct policy")][switch] $DisplayCorrectPolicy

###############################################
# Parameters validation
###############################################

#######################
# Validate Path
#######################

if (Test-Path -Path $Path -PathType Any)
{
    if (Test-Path -Path $Path -PathType Container)
    {
        $ADMXFilelist = Get-ChildItem -Path $Path -File -Filter "*.admx"
        if ( $ADMXFilelist.count -eq 0 )
        {
            Write-Error -Exception "List ADMX Files Error" -Message "NO ADMX file found, please specify a folder containing ADMX files." -Category InvalidArgument -TargetObject $Path -CategoryReason "Invalid Argument"
            exit
        }
    }
    else
    {
        $ADMXFilelist = Get-Item -Path $Path
        if ( $ADMXFilelist.Extension -ne ".admx")
        {
            Write-Error -Exception "File Type Error" -Message "Extension of the file is incorrect, please specify a valid ADMX file." -Category InvalidArgument -TargetObject $Path -CategoryReason "Invalid Argument"
            exit
        }
    }
}
else {
    Write-Error -Exception "Path Error" -Message "Path doesn't exist, please specify a valid folder or ADMX file." -Category InvalidArgument -TargetObject $Path -CategoryReason "Invalid Argument"
    exit
}

###############################################
# Internal Variables
###############################################

# List from : https://docs.microsoft.com/en-us/windows/client-management/mdm/win32-and-centennial-app-policy-configuration

#######################
# CSP Forbidden keys
#######################

$CSPForbiddenKeys = (
    'System',
    'Software\Microsoft',
    'Software\Policies\Microsoft'
)

#######################
# CSP Exception
#######################

$CSPExceptionLocation = (
    'Software\Policies\Microsoft\Office',
    'Software\Microsoft\Office',
    'Software\Microsoft\Windows\CurrentVersion\Explorer',
    'Software\Microsoft\Internet Explorer',
    'software\policies\microsoft\shared tools\proofing tools',
    'software\policies\microsoft\imejp',
    'software\policies\microsoft\ime\shared',
    'software\policies\microsoft\shared tools\graphics filters',
    'software\policies\microsoft\windows\currentversion\explore',
    'software\policies\microsoft\softwareprotectionplatform',
    'software\policies\microsoft\officesoftwareprotectionplatform',
    'software\policies\microsoft\windows\windows search\preferences',
    'software\policies\microsoft\exchange',
    'software\microsoft\shared tools\proofing tools',
    'software\microsoft\shared tools\graphics filters',
    'software\microsoft\windows\windows search\preferences',
    'software\microsoft\exchange',
    'software\policies\microsoft\vba\security',
    'software\microsoft\onedrive'
)

# Lower casing to avoid case sensitive issue
$CSPExceptionLocation = $CSPExceptionLocation.ToLowerInvariant()
$CSPForbiddenKeys = $CSPForbiddenKeys.ToLowerInvariant()


###############################################
# Functions
###############################################

#######################
# Name    : Test-PolicyKey
# Purpose : Test if the Policy Key is a valid key for CSP Import
# Input   : PolicyKey from ADMX
# Return  : Boolean : TRUE if valid / FALSE if NOT valid 
# Example : Test-PolicyKey -PolicyKey "software\Microsoft\Windows"
#######################
function Test-PolicyKey {
    param (
        [String]$PolicyKey
    )

    process {

        foreach ($ForbiddenKey in $CSPForbiddenKeys)
        {
            $RegExForbiddenKey = [RegEx]::Escape($ForbiddenKey)
            $PatternForbiddenKey = "\b$($RegExForbiddenKey)\w*\b"

            if ( $PolicyKey -imatch $PatternForbiddenKey )
            {
                #system don't have exclusion no need to test
                if  ($ForbiddenKey -ne "system" )
                {
                    #Checking if the Policy Key is in the exception list
                    foreach ($ValidKey in $CSPExceptionLocation)
                    {
                        $RegExValidKey = [RegEx]::Escape($ValidKey)
                        $PatternValidKey = "\b$($RegExValidKey)\w*\b"
                        if ( $PolicyKey -imatch $PatternValidKey )
                        {
                            return $true
                        }
                        Remove-Variable RegExValidKey,PatternValidKey
                    }
                    return $false
                }
                else
                {
                    #Policy Key Start with System, so it's not a valid key
                    return $false
                }
            }
            Remove-Variable RegExForbiddenKey,PatternForbiddenKey
        }
        return $true
    }
}

#######################
# Name    : Test-ADMX
# Purpose : Interate on each policy and call Test-PolicyKey for validation
# Input   : XML from ADMX import
# Return  : String : 
# Example : Test-PolicyKey -PolicyKey "software\Microsoft\Windows"
#######################
function Test-ADMX {
    param (
        [XML]$ADMXContent
    )

    process {
        $TestResult = New-Object -TypeName psobject
        $TestResult | Add-Member -MemberType NoteProperty -Name "ValidPolicy" -Value ""
        $TestResult | Add-Member -MemberType NoteProperty -Name "InvalidPolicy" -Value ""

        foreach ($Policy in $ADMXContent.policyDefinitions.policies.policy)
        {
            #############################################
            # Policy Example
            # name           : L_RestrictActiveXInstall
            # class          : Machine
            # displayName    : $(string.L_RestrictActiveXInstall)
            # explainText    : $(string.L_RestrictActiveXInstallExplain)
            # presentation   : $(presentation.L_RestrictActiveXInstall)
            # key            : software\microsoft\internet explorer\main\featurecontrol\feature_restrict_activexinstall
            # parentCategory : parentCategory
            # supportedOn    : supportedOn
            # elements       : elements
            #############################################

            if ( Test-PolicyKey -PolicyKey $Policy.Key )
            {
                # UnComment me to display valid CSP
                #Write-Host "Key $($Policy.Key) is valid for CSP" -ForegroundColor Green
                $TestResult.ValidPolicy += "Policy $($Policy.Name) is valid for CSP`r`n"
            }
            else 
            {
                #Write-Host "Policy $($Policy.Name) is not valid for CSP" -ForegroundColor Red -BackgroundColor White
                $TestResult.InvalidPolicy += "Policy $($Policy.Name) is not valid for CSP`r`n"
            }
        }

        return $TestResult
    }
}

#######################
# Test each ADMX
#######################

[String[]]$result
foreach ( $ADMX in $ADMXFilelist )
{
    try {
        Write-Host "---------------------------------------$($ADMX.Name)---------------------------------------" -BackgroundColor Black

        [XML]$ADMXFileContent = Get-Content -Path $ADMX.FullName
        $ADMXtestresult = Test-ADMX -ADMXContent $ADMXFileContent

        if ( $ADMXtestresult.InvalidPolicy -eq "")
        {   
            Write-Host "ADMX OK - No issue found" -ForegroundColor Green -BackgroundColor Black
            if ( $DisplayCorrectPolicy )
            {
                Write-Host $ADMXtestresult.ValidPolicy -ForegroundColor Green -BackgroundColor Black
            }
        }
        else 
        {
            Write-Host "Issue with ADMX Policy found" -ForegroundColor Red -BackgroundColor Black
            if ( $DisplayCorrectPolicy )
            {
                Write-Host $ADMXtestresult.ValidPolicy -ForegroundColor Green -BackgroundColor Black
            }
            Write-Host $ADMXtestresult.InvalidPolicy -ForegroundColor Red -BackgroundColor Black
        }
    }
    catch {
        Write-Error -Exception "ADMX import error" -Message "An error occured when importing ADMX file $($ADMX.Name), verify that the ADMX file is valid" -Category InvalidArgument -TargetObject $ADMX -CategoryReason "Invalid Argument"
    }
}