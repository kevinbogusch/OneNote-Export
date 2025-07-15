<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER

.EXAMPLE

.INPUTS

.OUTPUTS

.NOTES

.LINK
#>

# Validate parameters for the script itself
[CmdletBinding(DefaultParameterSetName = 'Set1')]
param(
    [Parameter(Position=0)]
    [string] $PageName,
    [Parameter(Mandatory=$False, ParameterSetName = 'Set1')]
    [switch] $Display,
    [Parameter(Mandatory=$False, ParameterSetName = 'Set2')]
    [switch] $Save,

    [Parameter(Mandatory=$False)]
    [switch] $PrintStructure,

    [Parameter(Mandatory=$False)]
    [switch] $v,
    [Parameter(Mandatory=$False)]
    [switch] $vv,
    [Parameter(Mandatory=$False)]
    [switch] $vvv,
    [Parameter(Mandatory=$False)]
    [switch] $vvvv

)

# Reference: Application interface (OneNote): https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote

# -----------------------------------------------------------------------------
# Types

# Because we're limited to Windows PowerShell 5, dictionaries do not preserve
# the order of insertions. This means that paragraphs within notes will not be
# generated in the order that they appear in the note. To work around this, we
# load the OrderedDictionary class from .NET. Note that the method 
# "ContainsKey" is replaced by the method "Contains". Otherwise, it operates
# in the same manner.

# It could be pointed out that the results will be made into files, which the
# OS will then sort in "correct" order, but we don't want to depend upon the OS.
Add-Type -TypeDefinition @"
using System.Collections.Specialized;
public class MyOrderedDict : OrderedDictionary {}
"@

# -----------------------------------------------------------------------------
# Reference values

$LogLevels = @("DEBUG", "VERBOSE", "INFO", "WARNING")
$LogLevel = ""

# -----------------------------------------------------------------------------

function Write-Log {
    param(
        [string]$Level,
        [string]$Message
    )

    switch ($Level) {
        "DEBUG" {
            If ($Loglevel -eq "DEBUG") {
                Write-Debug "$Message" 
            }
        }
        "VERBOSE" {
            If (($LogLevel -eq "DEBUG") -or ($LogLevel -eq "VERBOSE")) {
                Write-Host "VERBOSE: $Message" -ForegroundColor Blue
            }
        }
        "INFO" {
            If (($LogLevel -eq "DEBUG") -or ($LogLevel -eq "VERBOSE") -or ($LogLevel -eq "INFO")) {
                Write-Host "INFO: $Message" -ForegroundColor Green
            }
        }
        "WARNING" {
            If (($LogLevel -eq "DEBUG") -or ($LogLevel -eq "VERBOSE") -or ($LogLevel -eq "INFO") -or ($LogLevel -eq "WARNING")) {
                Write-Warning "$Message"
            }
        }
        "ERROR" {
            Write-Error "$Message"
        }	
        default {
        }
    }
    
}

# ==================================================================================================
# Main 

# If there is any logging level, then allow debug printing.
If ($v.IsPresent -or $vv.IsPresent -or $vvv.IsPresent -or $vvvv.IsPresent){
    $DebugPreference = "Continue"
    $WarningPreference = "Continue"
}

# Set logging level. In case multiple levels are specified, the highest level is used.
If ( $vvvv.IsPresent){
    $LogLevel = "DEBUG"
    Write-Log "INFO" "Log level set to DEBUG"
} ElseIf ( $vvv.IsPresent) {
    $LogLevel = "VERBOSE"
    Write-Log "INFO" "Log level set to VERBOSE"
} ElseIf ( $vv.IsPresent) {
    $LogLevel = "INFO"
    Write-Log "INFO" "Log level set to INFO"
} ElseIf ( $v.IsPresent) {
    # If LogLevel is set to WARNING, we can't print an INFO message that it is set to WARNING.
    $LogLevel = "WARNING"
} Else {
    # Default to logging only errors.
    $LogLevel = "ERROR"
}

Write-Log -Level "VERBOSE" -Message "PageName: $PageName"
Write-Log -Level "VERBOSE" -Message "Display: $Display"
Write-Log -Level "VERBOSE" -Message "Save: $Save"
Write-Log -Level "VERBOSE" -Message "v: $v"
Write-Log -Level "VERBOSE" -Message "vv: $vv"
Write-Log -Level "VERBOSE" -Message "vvv: $vvv"
Write-Log -Level "VERBOSE" -Message "vvvv: $vvvv"

# Advanced logic for parameters.
If ($PSCmdlet.ParameterSetName -eq "Set1") {
    $Display=$true
    $Save=$false
} ElseIf ($PSCmdlet.ParameterSetName -eq "Set2") {
    $Display=$false
    $Save=$true
}

$ILLEGAL_CHARACTERS = "[\\\/\:\*\?\`"\<\>\|]" 
Write-Log "DEBUG" "Illegal characters: `"$ILLEGAL_CHARACTERS`""
$IllegalCharactersInHex = ($PotentialIllegalCharacters | ForEach-Object { "{0:X}" -f [int]$_ }) -join ","
Write-Log "DEBUG" "Illegal characters in hex: `"$IllegalCharactersInHex`""

# Start the OneNote application COM object.

    # The following code can be used to manually load the assembly, but it *should* work without it.
    # $AssemblyFile = (get-childitem $env:windir\assembly -Recurse Microsoft.Office.Interop.OneNote.dll | Sort-Object Directory -Descending | Select-Object -first 1).FullName
    # Add-Type -Path $AssemblyFile -IgnoreWarnings

$OneNoteApp = New-Object -ComObject OneNote.Application

# Ask OneNote for the hiearchy all the way down to the individual pages, which is as low as you can
# go.
# The 2013 schema is the most recent schema available.
# If the notebook directory is not specified, then the default notebook is used.
Write-Log "DEBUG" "Getting the hierarchy of notebooks and pages."
[xml]$NotebooksXML = ""
$Scope = [Microsoft.Office.Interop.OneNote.HierarchyScope]::hsPages
$OneNoteVersion = [Microsoft.Office.Interop.OneNote.XMLSchema]::xs2013
try {
    $OneNoteApp.GetHierarchy($NotebookDir, $Scope, [ref]$NotebooksXML, $OneNoteVersion)
}
catch {
    Write-Log "ERROR" "An error occurred while getting the hierarchy of notebooks and pages."
    Exit 1
}
Write-Log "DEBUG" "Got the hierarchy of notebooks and pages."

$FoundPage = $false
ForEach($Notebook in $NotebooksXML.Notebooks.Notebook)
{
    If ($FoundPage) {
        break
    }
    
    Write-Log "VERBOSE" "Notebook Name: `"$($Notebook.Name.trim())`""
    If ($PrintStructure.IsPresent) {
        Write-Output "Notebook Name: ""$($Notebook.Name.trim())"""
    }

    ForEach($Section in $Notebook.Section)
    {
        If ($FoundPage) {
            break
        }
        Write-Log "VERBOSE" "Section Name: `"$($Section.Name.trim())`""
        If ($PrintStructure.IsPresent){
            Write-Output "- Section Name: ""$($Section.Name.trim())"""
        }

        ForEach($Page in $Section.Page) 
        {
            Write-Log "VERBOSE" "Page Name: `"$($Page.Name)`""
            If ($PrintStructure.IsPresent) {
                Write-Output "  - Page Name: ""$($Page.Name)"""
            }

            If ($Page.Name.trim().ToLower() -eq $PageName.ToLower()) {

                # This operation can potentially take a long time because it's fetching the entire 
                # contents of the page.
                Write-Log("DEBUG", "Getting the content of the page: `"$($Page.Name)`"")
                [xml]$PageXML = ""
                $OneNoteApp.GetPageContent($Page.ID, [ref]$PageXML, [Microsoft.Office.Interop.OneNote.PageInfo]::piBasic, $OneNoteVersion)
                Write-Log "DEBUG" "Got the content of the page: `"$($Page.Name)`""

                If ($Display){
                    Write-Output $PageXML.OuterXml
                }
                ElseIf ($Save) {
                    $CleansedPageName = $Page.Name.trim() -replace $ILLEGAL_CHARACTERS, "_"

                    If ($CleansedPageName -ne $Page.Name.trim()) {
                        Write-Log "VERBOSE" "The page name contains illegal characters. It has been cleansed to: `"$($CleansedPageName)`""
                    }
                    $PagePath = "$($CleansedPageName).xml"
                    Write-Log "DEBUG" "Saving the page to: `"$($PagePath)`""
                    $PageXML.OuterXml | Out-File -FilePath $PagePath -Encoding utf8
                }
                Else {
                    Write-Log "ERROR" "Neither Display nor Save was specified. Exiting."
                    Exit 1
                }
                $FoundPage = $true
                break
            }
            
        }
    }
}

# Clean up after yourself.
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($OneNoteApp) | Out-Null