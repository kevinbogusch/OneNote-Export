<#
.SYNOPSIS

.DESCRIPTION

.PARAMETER
- See https://norspire.atlassian.net/wiki/spaces/One/pages/101646374/Exporting+OneNote+pages#Parameters

.EXAMPLE

.INPUTS

.OUTPUTS

.NOTES
- Export_Notebooks does not validate the e_structure_ of the notes. It will faithfully reproduce the structure of the note into Markdown, such as starting with a second-level bullet without having a first-level bullet. This ig the garbage-in, garbage-out maxim.
- Export_Notebooks does not understand all HTML tags. It supports bold, italics, bold+italics, strikethrough, and anchor (links). All other tags are ignored.
- Export_Notebooks does not make use of all OneNote tags in the XML of a page nor their attributes. There is a loss of metadata when converting from OneNote pages into Markdown.

.LINK
#>

# Validate parameters for the script itself
[CmdletBinding(DefaultParameterSetName = 'Set1')]
param(
    [Parameter(Mandatory=$False, ParameterSetName = 'Set1')]
    [switch] $NoExport,
    [Parameter(Mandatory=$False, ParameterSetName = 'Set2')]
    [string] $ExportSelected,
    [Parameter(Mandatory=$False, ParameterSetName = 'Set3')]
    [switch] $ExportAll,

    [Parameter(Mandatory=$False)]
    [string] $NotebookDir,

    # The following three options are mutually exclusive. $NoPrintPage overrides $PrintSnippet overrides $PrintPage. $PrintSnippet is the default.
    [Parameter(Mandatory=$False)]
    [switch] $NoPrintPage,
    [Parameter(Mandatory=$False)]
    [switch] $PrintSnippet,
    [Parameter(Mandatory=$False)]
    [switch] $PrintPage,

    [Parameter(Mandatory=$False)]
    [string] $WhichPage,

    # The following three options are mutually exclusive. $Markdown overrides $PlainText overrides $HTML. $Markdown is the default.
    [Parameter(Mandatory=$False)]
    [switch] $Markdown,
    [Parameter(Mandatory=$False)]
    [switch] $PlainText,
    [Parameter(Mandatory=$False)]
    [switch] $HTML,

    [Parameter(Mandatory=$False)]
    [switch] $PrintStructure,
    [Parameter(Mandatory=$False)]
    [switch] $PrintStyles,
    [Parameter(Mandatory=$False)]
    [switch] $PrintTags,
    [Parameter(Mandatory=$False)]
    [switch] $SuppressOneNoteLinks,
    [Parameter(Mandatory=$False)]
    [switch] $NoDirCreation,
    [Parameter(Mandatory=$False)]
    [string] $ExportDir,

    [Parameter(Mandatory=$False)]
    [switch] $v,
    [Parameter(Mandatory=$False)]
    [switch] $vv,
    [Parameter(Mandatory=$False)]
    [switch] $vvv,
    [Parameter(Mandatory=$False)]
    [switch] $vvvv,

    [Parameter(Mandatory=$False)]
    [int] $vDelay

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

# -----------------------------------------------------------------------------

function Find-OneNoteOutline {
    param (
        $PageNode,     # Accept XmlElement or any XML node type
        [int]$Depth = 0
        )
        
        if (($PageNode.Name) -eq "one:Outline"){
            return $PageNode
        } else {
            # Recurse into child nodes (if any)
            foreach ($Child in $PageNode.ChildNodes) {
                $FoundNode = Find-OneNoteOutline -PageNode $Child -Depth ($Depth + 1)
                If ($FoundNode) {
                    return $FoundNode
                }
            }
        
        }
    }
    
# -----------------------------------------------------------------------------

function Get-Tags {
    param (
        $PageNode     # We're expecting the DocumentElement.
    )

    # Known page styles
    #   - To Do
    #   - Important
    #   - Question
    #   - Idea
    #   - Remember for later

    $FoundTags = @{}

    # We depend on the tags being identified by increasing index number, starting at zero. This lets
    # us add them sequentially to the array.
    # It's valid to have NO tags defined for a page.
    # Tags can also have a highlight color that applies to the entire paragraph. We ignore it.

    ForEach ($Child in $PageNode.ChildNodes){
        If ($Child.LocalName -eq "TagDef" ){
            If ($Child.GetAttribute("index")) {
                If ($Child.GetAttribute("name")){
                    # Write-Output "  + Page style name: $($Child.GetAttribute("name"))"
                    $FoundTags[$Child.GetAttribute("index")] = $Child.GetAttribute("name")
                }
            }
        }
    }

    return $FoundTags
}
# -----------------------------------------------------------------------------
function Get-PageStyles {
    param (
        $PageNode     # We're expecting the DocumentElement.
    )

    # Known page styles
    #   - PageTitle
    #   - h1, h2, h3, h4, h5, h6 (headings)
    #   - p (normal)
    #   - code
    #   - blockquote (quote)
    #   - cite (citation)

    $FoundStyles = @{}

    # We depend on the styles being identified by increasing index number, starting at zero. This
    # lets us add them sequentially to the array.

    ForEach ($Child in $PageNode.ChildNodes){
        If ($Child.LocalName -eq "QuickStyleDef" ){
            If ($Child.GetAttribute("index")) {
                If ($Child.GetAttribute("name")){
                    # Write-Output "  + Page style name: $($Child.GetAttribute("name"))"
                    $FoundStyles[$Child.GetAttribute("index")] = $Child.GetAttribute("name")
                }
            }
        }
    }

    return $FoundStyles
}

# -----------------------------------------------------------------------------

function FormatHTMLTo-Markdown {
    param(
        [string] $Text,
        [bool] $RemoveNewlines = $True
    )
    [string] $ReturnText = ""
    
    Write-Log "DEBUG" "Calling FormatHTMLTo-Markdown"
    write-Log "DEBUG" "Text: $($Text.Substring(0, [Math]::Min($Text.Length, 80)))"
    If ($RemoveNewlines){
        $Text = $Text -replace "(`r`n|`r|`n)", ""
        Write-Log "DEBUG" "Text after removing newlines: $($Text.Substring(0, [Math]::Min($Text.Length, 80)))"
    }

    $OpeningTagRegex = "^([^<]*?)(<[^/][^>]+?>)(.*)$"
    If ($Text -match $OpeningTagRegex){
        write-Log "DEBUG" "Opening tag found (first time)"
        do {
            If ($vDelay){
                Start-Sleep -Milliseconds $vDelay
            }

            $OpeningText = $matches[1]
            $OpeningTag = $matches[2]
            $RemainingOpeningText = $matches[3]
            Write-Log "DEBUG" "`$OpeningText: `"$($OpeningText.Substring(0, [Math]::Min($OpeningText.Length, 80)))`""
            Write-Log "DEBUG" "`$OpeningTag: `"$($OpeningTag.Substring(0, [Math]::Min($OpeningTag.Length, 80)))`""
            Write-Log "DEBUG" "`$RemainingOpeningText: `"$($RemainingOpeningText.Substring(0, [Math]::Min($RemainingOpeningText.Length, 80)))`""

            # Break tags don't have matching closing tags, so just consume them and don't look for
            # a closing tag (i.e., "</br>").
            If ($OpeningTag.ToLower() -ne "<br>"){
                [string] $ConvertedText = ""
                $ConvertedText = FormatHTMLTo-Markdown -Text $RemainingOpeningText -RemoveNewlines $False
                
                $ClosingTagRegex = "^([^<]*?)(</[^>]+?>)(.*)$"
                If ($ConvertedText -match $ClosingTagRegex) {
                    $BeforeClosingText = $matches[1]
                    # $ClosingTag = $matches[2]
                    $RemainingClosingText = $matches[3]
                }
                
                # Bold tags: <span style='font-weight:bold'> ... </span>
                # Italic tags: <span style='font-style:italic'> ... </span>
                # Links: <a href="URL">URL-name</a>
                # Strikethrough: <span style='text-decoration:line-through'> ... </span>
                #   - Strikethrough isn't universally supported. It's a markdown extension.

                $BoldRegex = "font-weight:[\s]*?bold"
                $ItalicRegex = "font-style:[\s]*?italic"
                $LinkRegex = "(<a[\s]*?href="")(.*)("">)"
                $StrikethroughRegex = "style='text-decoration:line-through'"

                If (($OpeningTag -match $BoldRegex) -and ($OpeningTag -match $ItalicRegex)) {
                    Write-Log "DEBUG" "Bold and italic tags found"
                    $ReturnText += $OpeningText + "___" + $BeforeClosingText + "___"

                } ElseIf ($OpeningTag -match $BoldRegex) {
                    Write-Log "DEBUG" "Bold tag found"
                    $ReturnText += $OpeningText + "__" + $BeforeClosingText + "__"

                } ElseIf ($OpeningTag -match $ItalicRegex) {
                    Write-Log "DEBUG" "Italic tag found"
                    $ReturnText += $OpeningText + "_" + $BeforeClosingText + "_"

                } ElseIf ($OpeningTag -match $LinkRegex) {
                    Write-Log "DEBUG" "Link tag found"
                    $ReturnText += $OpeningText + "[" + $BeforeClosingText + "]" + "(" + $matches[2] + ")"

                } ElseIf ($OpeningTag -match $StrikethroughRegex){
                    Write-Log "DEBUG" "Strikethrough tag found"
                    $ReturnText += $OpeningText + "~~" + $BeforeClosingText + "~~"

                } Else {
                    # We don't care about any other tags, so remove them
                    Write-Log "DEBUG" "Some other tag found; ignoring"
                    $ReturnText += $OpeningText + $BeforeClosingText
                }

                $Text = $RemainingClosingText
            } Else {
                Write-Log "DEBUG" "Break tag found"
                $Text = $RemainingOpeningText
            }

            Write-Log "DEBUG" "Text after processing: $($Text.Substring(0, [Math]::Min($Text.Length, 80)))"
        } while ($Text -match $OpeningTagRegex)

        $ReturnText += $Text

    } Else {
        Write-Log "DEBUG" "No opening tag found"
        $ReturnText = $Text
    }

    Write-Log "DEBUG" "ReturnText: $($ReturnText.Substring(0, [Math]::Min($ReturnText.Length, 80)))"
    return $ReturnText
}

# -----------------------------------------------------------------------------

function Get-Email {
    param(
        [string]$emailFilePath,
        [string]$spacingLeader = ""
    )
    [string]$EmailMessage = ""

    # Create an Outlook application instance
    $OutlookApp = New-Object -ComObject Outlook.Application

    # Open the email file
    $email = $OutlookApp.CreateItemFromTemplate($emailFilePath)

    # Extract email details
    $subject = $email.Subject
    $body = $email.Body
    $sender = $email.SenderName

    Write-Log "VERBOSE" "Email subject: `"$subject`" from `"$sender`""

    # Output the email details
    $EmailMessage += $spacingLeader + "--BEGIN EMAIL MESSAGE-------------------`n"
    $EmailMessage += $spacingLeader + "Subject: $subject`n"
    $EmailMessage += $spacingLeader + "Sender: $sender`n"
    $EmailMessage += $spacingLeader + "Body: $body`n"

    # Check for attachments
    if ($email.Attachments.Count -gt 0) {
        foreach ($attachment in $email.Attachments) {
            $EmailMessage += $spacingLeader + "  - Attachment: `"$($attachment.FileName)`"`n"
        }
    }
    $EmailMessage += $spacingLeader + "--END EMAIL MESSAGE---------------------`n"

    # Clean up
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($email) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutlookApp) | Out-Null

    return $EmailMessage
}

# -----------------------------------------------------------------------------
# Count-Tags is a recursive function that will descende the XML tree and count every tag without
# processing it.

function Count-Tags {
    param (
        $PageNode,     # Accept any XML node type
        $TagCount = 0
    )

    $TagCount += 1

    ForEach ($Child in $PageNode.ChildNodes){
        $TagCount = Count-Tags -PageNode $Child -TagCount $TagCount
    }

    return $TagCount
}

# -----------------------------------------------------------------------------

# Convert-Page is a recursive function that will descend the XML tree of a OneNote page and convert 
# it to Markdown.
function Convert-Page {
    param (
        $PageName,
        $PageID,           # Object ID of the page
        $PageNode,         # Accept any XML node type
        $PageStyles,       # Hashtable of instantiated page styles {index, style_name}
        $Tags,             # Hashtable of instantiated tags {index, tag_name}
        $LastObjectID,     # Object ID of the parent
        $LastStyleName,    # StyleName of parent
        $IndentLevel = -1, # We start at -1 because the top tag is already an OEChildren
        $BulletLevel = 0,
        $AllNodeCount = 0, # The total number of nodes in the XML document
        $SumNodeCount = 0, # The sum of all nodes traversed
        [int]$ProgressCounterDelay = 0,
        [int]$ProgressCounter = 0
        )

        [String]$Paragraph = ""
        [int]$Bullet = $BulletLevel
        $ObjectID = $LastObjectID
        
    If ($vDelay){
        Start-Sleep -Milliseconds $vDelay
    }

    If ($AllNodeCount -eq 0){
        $AllNodeCount = Count-Tags $PageNode
        Write-Log "DEBUG" "AllNodeCount: $AllNodeCount"
    }

    # If the ProgressCounterDelay is zero, then set it to 5% of the total number of nodes. 5% is a 
    # compromise between updating the progress bar too frequently and not frequently enough.
    If ($ProgressCounterDelay -eq 0){
        Write-Log "DEBUG" "Setting `$ProgressCounterDelay: $ProgressCounterDelay"
        $ProgressCounterDelay = [int]($AllNodeCount * 0.05)
    }
        
    $SumNodeCount += 1
    Write-Log "DEBUG" "Node: $($PageNode.Name), SumNodeCount: $SumNodeCount"

    If (($Loglevel -eq "INFO") -or ($LogLevel -eq "VERBOSE") -or ($LogLevel -eq "DEBUG")){
        Write-Log "DEBUG" "`$SumNodeCount: $SumNodeCount, `$AllNodeCount: $AllNodeCount, `$ProgressCounter: $ProgressCounter, `$ProgressCounterDelay: $ProgressCounterDelay"
        $ProgressCounter += 1
        If ($ProgressCounter -gt $ProgressCounterDelay){
            Write-Progress -ID 1 -Activity "$($PageName)" -Status "Converting page" -PercentComplete (($SumNodeCount / $AllNodeCount) * 100)
            Write-Log "DEBUG" "Converting page progress: $(($SumNodeCount / $AllNodeCount) * 100)"
            $ProgressCounter = 0
        }
    }

    $StyleName = $LastStyleName
    If ($PageNode.quickStyleIndex){
        $StyleName = $PageStyles[$PageNode.quickStyleIndex]
        Write-Log "DEBUG" "StyleName: $StyleName"
    } 

    If ($PageNode.Name -eq "one:OEChildren"){
        $IndentLevel += 1

    } ElseIf ($PageNode.Name -eq "one:OE"){
        $ObjectID = $PageNode.objectID

    } ElseIf ($PageNode.Name -eq "one:Image"){
        Write-Log "DEBUG" "Image found"

        # Only include the alt text of the image if it is available.
        If ($PageNode.alt){
            $Paragraph = "{Image: ""$($PageNode.alt)""}`n"
        }
        
    } ElseIf ($PageNode.Name -eq "one:InsertedFile" ) {
        Write-Log "DEBUG" "Inserted file found"

        # Only include the file name and the original location of the file if they are available.
        If ($PageNode.preferredName){
            $Paragraph = "{File: ""$($PageNode.preferredName)"""
            If ($PageNode.pathSource) {
                $Paragraph += ", originally located at ""$($PageNode.pathSource)"""
            }
            $Paragraph += "}`n"

            If ( [System.Io.Path]::GetExtension($PageNode.preferredName) -eq ".msg" ){
                # Email message
                If (Test-Path $PageNode.pathCache) {
                    $EmailMessage = Get-Email -emailFilePath $PageNode.pathCache -spacingLeader "    "
                    If ($EmailMessage) {
                        $Paragraph += $EmailMessage
                    }
                }
            }
        }

    } ElseIf ($PageNode.Name -eq "one:Bullet"){
        Write-Log "DEBUG" "Bullet found: $Bullet"
        $Bullet = $PageNode.GetAttribute("bullet")

    } ElseIf ($PageNode -is [System.Xml.XmlCDataSection]){
        # Only actual text that is typed into OneNote appears in ![CDATA] sections.
        Write-Log "DEBUG" "CDATA found"

        If ($PageNode.Value.trim() -ne "") {
            Write-Log "VERBOSE" "CDATA: $($PageNode.Value.trim().Substring(0, [Math]::Min($PageNode.Value.trim().Length, 40)))"
            If ($Markdown.IsPresent){
                # We'll replace bold, italics, strikethrough, and links.
                # All other HTML tags will be removed. This includes but isn't limited to:
                #   - Font name 
                #   - Font size
                #   - 

                Write-Log "DEBUG" "Converting CDATA to Markdown"
                $PageText = FormatHTMLTo-Markdown -Text $PageNode.Value.trim()

            } ElseIf ($HTML.IsPresent) {
                Write-Log "DEBUG" "Leaving CDATA as HTML"
                $PageText = $PageNode.Value.trim()

            } Else { 
                # Default to $PlainText=$True
                Write-Log "DEBUG" "Remove HTML from CDATA as plain text"
                $PageText = $PageNode.Value.trim() -replace "<[^>]+>", ""
            }

            Write-Log "DEBUG" "PageText processed: $($PageText.Substring(0, [Math]::Min($PageText.Length, 40)))"
            $Leader = ""
            Write-Log "DEBUG" "StyleName: $StyleName"
            Switch ($StyleName){
                "PageTitle"{
                    $Leader += ""
                }
                "h1" {
                    $Leader += "# "
                }
                "h2" {
                    $Leader += "## "
                }
                "h3" {
                    $Leader += "### "
                }
                "h4" {
                    $Leader += "#### "
                }
                "p" {
                    $Leader += ""
                }
                "blockquote" {
                    $Leader += "> "
                }
                "code" {
                    $Leader += "``` "
                }
                "cite" {
                    $Leader += ""
                }
                default {
                    $Leader += ""
                }
            }
            $Paragraph = "$($Leader)"
            If ($Bullet -gt 0){
                $Paragraph += " " * ((($Bullet) -1)* 2) + "+ "
                $Bullet = 0
            } Else {
                $Paragraph += " " * (($IndentLevel) * 2)
            }
            $Paragraph += "$($PageText) `n"

            If (-not $SuppressOneNoteLinks.IsPresent){
                If ($StyleName -eq "h1"){
                    $HyperlinkToObject = ""
                    Write-Log "DEBUG" "About to get hyperlink to object because it is an h1 header"
                    $OneNoteApp.GetHyperLinkToObject( $PageID, $ObjectID, [ref]$HyperlinkToObject)
                    Write-Log "DEBUG" "Hyperlink to object: $HyperlinkToObject.substring(0, [Math]::Min($HyperlinkToObject.Length, 40))"
                    $PlainPageText = $PageNode.Value.trim() -replace "<[^>]+>", ""
                    $Paragraph += "[$($PlainPageText)]($($HyperlinkToObject))`n"
                    $ObjectID = ""
                }
            }
        }
    }

    # Recurse into child nodes (if any)
    foreach ($Child in $PageNode.ChildNodes) {
        $ConvertResult = Convert-Page -PageName $PageName -PageID $PageID -PageNode $Child -PageStyles $PageStyles $LastObjectID $ObjectID -LastStyleName $StyleName -IndentLevel $IndentLevel -BulletLevel $Bullet -AllNodeCount $AllNodeCount -SumNodeCount $SumNodeCount -ProgressCounterDelay $ProgressCounterDelay -ProgressCounter $ProgressCounter
        $Paragraph += $ConvertResult.Paragraph
        $Bullet = $ConvertResult.Bullet
        $SumNodeCount = $ConvertResult.SumNodeCount
        $ProgressCounter = $ConvertResult.ProgressCounter
    }

    If (($Loglevel -eq "INFO") -or ($LogLevel -eq "VERBOSE") -or ($LogLevel -eq "DEBUG")){
        If ($AllNodeCount -eq $SumNodeCount){
            Write-Progress -ID 1 -Activity "$($PageName)" -Status "Converting page" -Completed
            Write-Log "DEBUG" "Converting page progress: Completed"
        }
    }

    return [PSCustomObject]@{
        Paragraph = $Paragraph
        Bullet = $Bullet
        SumNodeCount = $SumNodeCount
        ProgressCounter = $ProgressCounter
    }

}

# -----------------------------------------------------------------------------
function Split-Pages {
    param (
        $PageMarkdown,
        $PageName,
        $AllH1Count = 0,
        $SumH1Count = 0,
        [int]$ProgressCounterDelay = 0
    )

    # $PageParagraphs = @{}
    $PageParagraphs = New-Object MyOrderedDict

    If ($AllH1Count -eq 0){
        $AllH1Count = ($PageMarkdown -split "`r?`n" | Select-String "^# ").Count
        Write-Log "DEBUG" "AllH1Count: $AllH1Count"
    }

    # If the ProgressCounterDelay is zero, then set it to 5% of the total number of nodes. 5% is a 
    # compromise between updating the progress bar too frequently and not frequently enough.
    If ($ProgressCounterDelay -eq 0){
        $ProgressCounterDelay = [int]($AllNodeCount * 0.05)
    }
    $ProgressCounter = 0
    
    $Lines = $PageMarkdown -split "`r?`n"
    $LastParagraphTitle = ""
    $LastParagraph = ""
    
    ForEach($Line in $Lines){
        If ( $vDelay){
            Start-Sleep -Milliseconds $vDelay
        }

        If ($($Line) -match "^# ") {
            $Line = $Line.TrimEnd()
            Write-Log "VERBOSE" "H1 heading: `"$($Line)`""
            
            $SumH1Count += 1
            If (($Loglevel -eq "INFO") -or ($LogLevel -eq "VERBOSE") -or ($LogLevel -eq "DEBUG")){
                $ProgressCounter += 1
                If ($ProgressCounter -gt $ProgressCounterDelay){
                    Write-Progress -ID 2 -Activity "$($PageName)" -Status "Splitting pages" -PercentComplete (($SumH1Count / $AllH1Count) * 100)
                    Write-Log "DEBUG" "Splitting pages progress: $(($SumH1Count / $AllH1Count) * 100)"
                    $ProgressCounter = 0
                }
            }

            If ($PageParagraphs.Contains($LastParagraphTitle)){
                $PageParagraphs[$LastParagraphTitle] += "$($LastParagraph)"
                
            } Else {
                If ($LastParagraphTitle -eq ""){
                    If ($LastParagraph){
                        $LastParagraphTitle = $PageName
                        $PageParagraphs.Add($LastParagraphTitle, "$($LastParagraph)")
                    }
                } Else {
                    $PageParagraphs.Add($LastParagraphTitle, "$($LastParagraph)")
                }
            }
            
           $LastParagraph=""

            # [1] is the leading hash mark, [2] is the leading whitespace and potential bullet(s),
            # and [3] is the title.
            $TitleLine = $Line -match "^(# )([\s]*?[\*+-]*[\s]*)(.*)$"
            $NoDateMatch = "$($Matches[2])$($Matches[3])"
            If ($matches[3]){
                # There is NO guarantee that what we have is a date. NONE. So,
                # we'll try to convert it a few different ways. If that's not
                # successful, then we'll just use it as-is.
                $LastParagraphTitle = $Matches[3]
                
                # First, let's assume the typical date format of
                # "dddd, MMMM dd, yyyy" (e.g. Friday, January 01, 2027).

                # + Remove the leading day name, if any, from the expected
                #   format of "dddd, MMMM dd, yyyy". Occasionally, the wrong
                #   day is attached to the right date. 
                $CleanedDate = $LastParagraphTitle -replace "^[^,]+,\s*", ""
                Write-Log "DEBUG" "Cleaned date (dddd, MMMM dd, yyyy): `"$CleanedDate`""
                
                # + If we can recognize the heading as a valid date in the form
                #   of "MMMM dd, yyyy", then reformat it to "yyyy-MM-dd", which
                #   is friendlier to a file system.
                # + Also add the full date ("dddd, MMMM dd, yyyy") to the cache
                #   so that it is available as text to the LLM.
                try{
                    If ((Get-Date $CleanedDate).ToString("yyyy-MM-dd")){
                        $LastParagraphTitle = (Get-Date $CleanedDate).ToString("yyyy-MM-dd")
                    } 
                    Write-Log "VERBOSE" "Date recognized: `"$LastParagraphTitle`""
                }
                catch {
                    # Second, let's assume the alternate date format of
                    # "MMMM dd, yyyy, dddd" (e.g. January 01, 2027, Friday).

                    $CleanedDate = $LastParagraphTitle -replace ",\s*[^,]+$", ""
                    Write-Log "DEBUG" "Cleaned date (MMMM dd, yyyy, dddd): `"$CleanedDate`""
                    try{
                        If ((Get-Date $CleanedDate).ToString("yyyy-MM-dd")){
                            $LastParagraphTitle = (Get-Date $CleanedDate).ToString("yyyy-MM-dd")
                        } 
                        Write-Log "VERBOSE" "Date recognized: `"$LastParagraphTitle`""
                    }
                    catch {
                        # At this point, it doesn't match either of the date
                        # formats that we expect, so just use it as-is.
                        $LastParagraphTitle = $NoDateMatch
                        Write-Log "DEBUG" "Date not recognized: `"$NoDateMatch`""
                    }
                }
            } Else {
                # + We know that the line is an h1 heading (starts with "^# ")
                #   but it doesn't match the regular expression above. So, just
                #   set the line value to whatever comes after the "^# ".
                # + Note that the "Untitled" section can only occur once, at the
                #   beginning of the page. After any h1 heading has been 
                #   encountered, a section will *always* have a name.
                $LastParagraphTitle = $NoDateMatch
            }

            Write-Log "DEBUG" "Determined `$LastParagraphTitle: `"$LastParagraphTitle`""

            If (!($PageParagraphs.Contains($LastParagraphTitle))) {
                try {
                    $LastParagraph = "# " + (Get-Date $LastParagraphTitle).ToString("dddd, MMMM dd, yyyy") +"`n"
                } 
                catch {
                    $LastParagraph = "# $($LastParagraphTitle)" + "`n"
                }
            }

        } Else {
            # + The line is not an h1 heading. Just add it to the cache.
            $LastParagraph += "$($Line)`n"
        }
    }

    # Issue-3
    # + We only add lines when we encounter an h1 heading. However, if an h1 
    #   heading is *not* the last line, we'll leave lines in the $LastParagraph
    #   buffer without adding them. So, look in $LastParagraph and see if any
    #   lines need to be added.
    If ($LastParagraph) {
        If ($PageParagraphs.Contains($LastParagraphTitle)){
            $PageParagraphs[$LastParagraphTitle] += "$($LastParagraph)"
        } Else {
            If ($LastParagraphTitle -eq ""){
                If ($LastParagraph){
                    $LastParagraphTitle = $PageName
                    $PageParagraphs.Add($LastParagraphTitle, "$($LastParagraph)")
                }
            } Else {
                $PageParagraphs.Add($LastParagraphTitle, "$($LastParagraph)")
            }        }
    }

    If (($Loglevel -eq "INFO") -or ($LogLevel -eq "VERBOSE") -or ($LogLevel -eq "DEBUG")){
        If ($AllNodeCount -eq $SumNodeCount){
            Write-Progress -ID 2 -Activity "$($PageName)" -Status "Splitting pages" -Completed
            Write-Log "DEBUG" "Splitting pages progress: Completed"
        }
    }

    return $PageParagraphs
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

Write-Log -Level "VERBOSE" -Message "NoExport: $NoExport"
Write-Log -Level "VERBOSE" -Message "ExportSelected: $ExportSelected"
Write-Log -Level "VERBOSE" -Message "ExportAll: $ExportAll"
Write-Log -Level "VERBOSE" -Message "NotebookDir: $NotebookDir"
Write-Log -Level "VERBOSE" -Message "NoPrintPage: $NoPrintPage"
Write-Log -Level "VERBOSE" -Message "PrintSnippet: $PrintSnippet"
Write-Log -Level "VERBOSE" -Message "PrintPage: $PrintPage"
Write-Log -Level "VERBOSE" -Message "WhichPage: `"$WhichPage`""
Write-Log -Level "VERBOSE" -Message "Markdown: $Markdown"
Write-Log -Level "VERBOSE" -Message "PlainText: $PlainText"
Write-Log -Level "VERBOSE" -Message "HTML: $HTML"
Write-Log -Level "VERBOSE" -Message "PrintStructure: $PrintStructure"
Write-Log -Level "VERBOSE" -Message "PrintStyles: $PrintStyles"
Write-Log -Level "VERBOSE" -Message "PrintTags: $PrintTags"
Write-Log -Level "VERBOSE" -Message "SuppressOneNoteLinks: $SuppressOneNoteLinks"
Write-Log -Level "VERBOSE" -Message "NoDirCreation: $NoDirCreation"
Write-Log -Level "VERBOSE" -Message "ExportDir: `"$ExportDir`""
Write-Log -Level "VERBOSE" -Message "v: $v"
Write-Log -Level "VERBOSE" -Message "vv: $vv"
Write-Log -Level "VERBOSE" -Message "vvv: $vvv"
Write-Log -Level "VERBOSE" -Message "vvvv: $vvvv"
Write-Log -Level "VERBOSE" -Message "vDelay: $vDelay"

# Advanced logic for parameters.
If ($PSCmdlet.ParameterSetName -eq "Set1") {
    $NoExport=$True
}

# If no print option is specified, then default to PrintSnippet.
If ((-not $NoPrintPage.IsPresent) -and (-not $PrintPage.IsPresent)){
    Write-Log("INFO", "No print option specified. Defaulting to PrintSnippet.")
    $PrintSnippet=$True
} Else {
    # If $PrintSnippet is specified, ignore the other print options.
    If ($PrintSnippet.IsPresent){
        $NoPrintPage=$false
        $PrintPage=$false
    } ElseIf ($NoPrintPage.IsPresent){
        # If $NoPrint is specified, ignore the $PrintPage option (whether it was specified or not).
        # Also ignore the $PrintSnippet option for good hygiene.
        $PrintSnippet=$false
        $PrintPage=$false
    } Else {
        # $PrintPage must be set.
        $PrintPage=$True
        $PrintSnippet=$false
        $NoPrintPage=$false
    }
}
Write-Log "VERBOSE" "After parameter logic: NoPrintPage: $NoPrintPage"
Write-Log "VERBOSE" "After parameter logic: PrintSnippet: $PrintSnippet"
Write-Log "VERBOSE" "After parameter logic: PrintPage: $PrintPage"

If ((-not $HTML) -and (-not $PlainText)){
    Write-Log("INFO", "No output format specified. Defaulting to Markdown.")
    $Markdown=$True
} Else {
    # If $Markdown is specified, ignore the other output options.
    If ($Markdown){
        $HTML=$false
        $PlainText=$false
    } ElseIf ($PlainText){
        $Markdown=$false
        $HTML=$false
    } Else {
        # $HTML must be set.
        $HTML=$True
        $PlainText=$false
    }
}
Write-Log "VERBOSE" "After parameter logic: Markdown: $Markdown"
Write-Log "VERBOSE" "After parameter logic: PlainText: $PlainText"
Write-Log "VERBOSE" "After parameter logic: HTML: $HTML"

$ILLEGAL_CHARACTERS = "[\\\/\:\*\?\`"\<\>\|]" 
Write-Log "DEBUG" "Illegal characters: `"$ILLEGAL_CHARACTERS`""
$IllegalCharactersInHex = ($PotentialIllegalCharacters | ForEach-Object { "{0:X}" -f [int]$_ }) -join ","
Write-Log "DEBUG" "Illegal characters in hex: `"$IllegalCharactersInHex`""

# NoDirCreation overrides the ExportDir parameter, if ExportDir is specified.
If (!$NoDirCreation){
    If ($ExportDir){
        If (!(Test-Path -Path $ExportDir -PathType Container)) {
            Write-Log "ERROR" "The specified export directory does not exist."
            Exit 1
        }
    } Else {
        Write-Log("INFO", "No export directory specified. Defaulting to the current directory.")
        $ExportDir = "."
    }
}

# Ensure that the notebook directory is valid. A null value is acceptable.
If ($NotebookDir){
    If (!(Test-path $NotebookDir)){
        Write-Log "ERROR" "The specified notebook directory does not exist."
        Exit 1
    }
}

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

ForEach($Notebook in $NotebooksXML.Notebooks.Notebook)
{
    Write-Log "VERBOSE" "Notebook Name: `"$($Notebook.Name.trim())`""
    If ($PrintStructure.IsPresent) {
        Write-Output "Notebook Name: ""$($Notebook.Name.trim())"""
    }

    $CleansedNotebookName = $Notebook.Name.trim() -replace $ILLEGAL_CHARACTERS, "_"
    If ($CleansedNotebookName -ne $Notebook.Name.trim()){
        Write-Log "VERBOSE" "The notebook name contains illegal characters. It has been cleansed to: `"$($CleansedNotebookName)`""
    }
    $NotebookPath = Join-Path -Path $ExportDir -ChildPath "$($CleansedNotebookName) notebook"
    If ((!$NoExport) -and (!$NoDirCreation)) {
        If (!(Test-Path -Path $NotebookPath -PathType Container)) {
            Write-Log "INFO" "Creating the notebook directory: `"$($CleansedNotebookName) notebook`""
            New-Item -Path $NotebookPath -ItemType Directory | Out-Null
        }
    }

    ForEach($Section in $Notebook.Section)
    {
        Write-Log "VERBOSE" "Section Name: `"$($Section.Name.trim())`""
        If ($PrintStructure.IsPresent){
            Write-Output "- Section Name: ""$($Section.Name.trim())"""
        }

        $CleansedSectionName = $Section.Name.trim() -replace $ILLEGAL_CHARACTERS, "_"
        If ($CleansedSectionName -ne $Section.Name.trim()){
            Write-Log "VERBOSE" "The section name contains illegal characters. It has been cleansed to: `"$($CleansedSectionName)`""
        }
        $SectionPath = Join-Path -Path $NotebookPath -ChildPath "$($CleansedSectionName) section"
        If ((!$NoExport) -and (!$NoDirCreation)) {
            If ( !(Test-Path -Path $SectionPath -PathType Container)) {
                Write-Log "INFO" "Creating the section directory: `"$($CleansedSectionName) section`""
                New-Item -Path $SectionPath -ItemType Directory | Out-Null
            }
        }

        ForEach($Page in $Section.Page) 
        {
            Write-Log "VERBOSE" "Page Name: `"$($Page.Name)`""
            If ($PrintStructure.IsPresent) {
                Write-Output "  - Page Name: ""$($Page.Name)"""
            }

            If (($WhichPage.trim().Length -eq 0) -or ($WhichPage -eq $Page.Name.trim())) {
                Write-Log "VERBOSE" "The page name matches `$WhichPage parameter. Processing just this page."

                $CleansedPageName = $Page.Name.trim() -replace $ILLEGAL_CHARACTERS, "_"
                If ($CleansedPageName -ne $Page.Name.trim()){
                    Write-Log "VERBOSE" "The page name contains illegal characters. It has been cleansed to: `"$($CleansedPageName)`""
                }
                $PagePath = Join-Path -Path $SectionPath -ChildPath "$($CleansedPageName) page"
                If ((!$NoExport) -and (!$NoDirCreation)) {
                    If (!(Test-Path -Path $PagePath -PathType Container)){
                        Write-Log "INFO" "Creating the page directory: `"$($CleansedPageName) page`""
                        New-Item -Path $PagePath -ItemType Directory | Out-Null
                    }
                }

                # This operation can potentially take a long time because it's fetching the entire 
                # contents of the page.
                Write-Log("DEBUG", "Getting the content of the page: `"$($Page.Name)`"")
                [xml]$PageXML = ""
                $OneNoteApp.GetPageContent($Page.ID, [ref]$PageXML, [Microsoft.Office.Interop.OneNote.PageInfo]::piBasic, $OneNoteVersion)
                Write-Log "DEBUG" "Got the content of the page: `"$($Page.Name)`""
                
                $DelimiterPrinted = $False

                # It is possible that a page is empty, but even empty pages still have a defined style
                # for the page title. This is a good check for the validity of the page.

                # The styles can be uniquely defined for each page. Styles are
                # defined in the order that they are first used on the page. So,
                # h2 migh tbe style #2 or #7, depending on when it was first
                # used on that page.
                Write-Log "DEBUG" "Finding the styles of the page: `"$($Page.Name)`""
                $PageStyles = @{}
                $PageStyles = Get-PageStyles $PageXML.DocumentElement
                Write-Log "DEBUG" "Found the styles of the page: `"$($Page.Name)`""

                Write-Log "VERBOSE" "Styles found: $($PageStyles.Keys.Count)"
                If ( $PageStyles.Keys.Count -eq 0 ){
                    # All pages must have at least the PageTitle style.
                    Write-Log "ERROR" "Could not find any styles for the page $($Page.Name)"
                    Exit 1
                }
                
                If ($PrintStyles){
                    Write-Output " "
                    Write-Output "--------"
                    $DelimiterPrinted=$True
                    Write-Output "  + Page Styles"
                    ForEach( $Key in $PageStyles.Keys) {
                        Write-Output "    - $($Key): $($PageStyles[$Key])"
                    }
                    Write-Output " "
                }    

                # Start at the beginning of the content, which is the first Outline node. There
                # can be multiple Outline nodes, but we're only interested in the first one.
                Write-Log "DEBUG" "Finding the first outline node of the page: `"$($Page.Name)`""
                [System.Xml.XmlElement]$OneNoteOutline = Find-OneNoteOutline $PageXML.DocumentElement 6
                
                If (!($OneNoteOutline)){
                    Write-Log "VERBOSE" "Could not find the first outline node of the page: `"$($Page.Name)`". The page is considered empty and will be ignored."
                } Else {
                    Write-Log "DEBUG" "Found the first outline node of the page: `"$($Page.Name)`""
                    
                    # Tags are optional. There might not be any on the page.
                    # It's possible that a page is empty and it's not an error to check for tags, but an
                    # empty page will never have tags so we don't check unless the page has some
                    # content. This is a performance optimization.
                    Write-Log "DEBUG" "Finding the tags of the page: `"$($Page.Name)`""
                    $Tags = @{}
                    $Tags = Get-Tags $PageXML.DocumentElement
                    Write-Log "DEBUG" "Found the tags of the page: `"$($Page.Name)`""
                    
                    Write-Log "VERBOSE" "Tags found: $($Tags.Keys.Count)"
                    If ($PrintTags){
                        If ($Tags.Keys.Count -gt 0){
                            If (!($DelimiterPrinted)){
                                Write-Output " "
                                Write-Output "--------"
                                $DelimiterPrinted=$True
                            }
                            Write-Output "  + Tags"
                            ForEach( $Key in $Tags.Keys) {
                                Write-Output "    - $($Key): $($Tags[$Key])"
                            }
                            Write-Output " "
                        }
                    }     
        
                    
                    If (($PrintSnippet) -or ($PrintPage)){
                        If (!($DelimiterPrinted)){
                            Write-Output " "
                            Write-Output "--------"
                            $DelimiterPrinted=$True
                        }
                    }
                        
                    # Convert the page to Markdown
                    Write-Log "DEBUG" "Converting the page from XML: `"$($Page.Name)`""
                    $ConvertResult = Convert-Page $Page.Name $Page.id $OneNoteOutline $PageStyles ""
                    Write-Log "DEBUG" "Converted the page from XML: `"$($Page.Name)`""

                    # Split the page into individual paragraphs and then
                    # write them to individual files.
                    Write-Log "DEBUG" "Splitting the page into paragraphs: `"$($Page.Name)`""
                    $PageParagraphs = Split-Pages -PageMarkdown $ConvertResult.Paragraph -PageName $Page.Name
                    Write-Log "DEBUG" "Split the page into paragraphs: `"$($Page.Name)`""

                    Write-Log "VERBOSE" "Paragraphs found: $($PageParagraphs.Keys.Count)"
                    ForEach ($PageParagraph in $PageParagraphs.Keys){
                        If ($PrintStructure.IsPresent) {
                            Write-Output "    * Paragraph Name: ""$($PageParagraph)"""
                        }

                        If ($PrintPage){
                            If (!($DelimiterPrinted)){
                                Write-Output " "
                                Write-Output "--------"
                                $DelimiterPrinted=$True
                            }
                            Write-Output("$($PageParagraphs[$PageParagraph])")
                        } ElseIf ($PrintSnippet){
                            If (!($DelimiterPrinted)){
                                Write-Output " "
                                Write-Output "--------"
                                $DelimiterPrinted=$True
                            }

                            #Write-Output($($PageParagraphs[$PageParagraph]) -split "`n" | Select-Object -First 3)
                            $Snippet = $($PageParagraphs[$PageParagraph]).Substring(0, [Math]::Min($($PageParagraphs[$PageParagraph]).Length, 100))
                            If ($Snippet.Length -eq 100){
                                $Snippet += "..."
                            }
                            Write-Output($Snippet)
                        }

                        $CleansedPageParagraph = $PageParagraph -replace $ILLEGAL_CHARACTERS, "_"
                        If ($CleansedPageParagraph -ne $PageParagraph){
                            Write-Log "VERBOSE" "The paragraph name `"$($PageParagraph)`" contains illegal characters. It has been changed to `"$($CleansedPageParagraph)`"."
                        }
                        
                        If (($ExportAll) -or (($ExportSelected) -and ($Page.Name -eq $ExportedSelected))){
                            # ONE-2
                            # Remove illegal characters from the paragraph name, which will then be the file name.
                            $PageParagraphFileName = Join-Path -Path $PagePath -ChildPath "$($CleansedPageParagraph.TrimEnd()).md"
                            $PageParagraphs[$PageParagraph].TrimEnd() | Out-File -FilePath $PageParagraphFileName -Encoding utf8
                        }
                    }

                    If ($DelimiterPrinted){
                        Write-Output "--------"
                        Write-Output " "
                    }
                }
            }
        }
    }
}

# Clean up after yourself.
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($OneNoteApp) | Out-Null