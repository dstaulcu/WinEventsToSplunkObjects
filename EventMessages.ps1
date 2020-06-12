$uri = "https://systemcenter.wiki/"

<#
$DebugPreference = "Continue"           # Debug Mode
$DebugPreference = "SilentlyContinue"   # Normal Mode
#>

function Get-ElementLinks {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("Task", "Discovery", "Rule","LinkedReport")]
    [String[]]
    $ElementType)

    $ElementUri = "$($uri)?GetElements=$($ElementType)&Category=$($Category)"
    $ElementContent = Invoke-WebRequest -Uri $ElementUri
    $Links = @($ElementContent.Links)
    # | ?{$_.innerText -match "$($ElementType)"})

    return $Links
}

# grab the main page of the uri
$WebContent = Invoke-WebRequest -Uri $uri

# get the link objects relating to Management Package Catalog
$mpCatalog = $WebContent.Links | ?{$_.outerHtml -match "GetCategory"}

# here is a key word list of products we care about
$ProductKeyWords = @("Active Directory","Cisco","Citrix","Dell Client","DFS","DFS","F5","File Services","Printer Monitoring","PowerShell","QLogic","RedHat","Skype","SQL","Docker","Operating System","Cluster","DHCP","DNS","iSCSI","Group Policy","IIS","Xerox")
$ProductKeyWords = @("Skype.*2019")
$ProductKeyWordsString = $ProductKeyWords -join "|"

# get the mp catalog link relating to products we care abouut
$mpCatalog = $mpCatalog | ?{$_.innerText -match "($($ProductKeyWordsString))"}

# 61 = windows client

$PerfmonObjects = @()
$WinEventLogObjects = @()

foreach ($item in $mpCatalog) {

    $result = Invoke-WebRequest -Uri $item.href 
    $resultTable = @{}

    # Get the title
    $resultTable.title = $result.ParsedHtml.title

    # Get the HTML Tag
    $HtmlTag = $result.ParsedHtml.childNodes | Where-Object {$_.nodename -eq 'HTML'} 

    # Get the HEAD Tag
    $HeadTag = $HtmlTag.childNodes | Where-Object {$_.nodename -eq 'HEAD'}

    # Get the Meta Tags
    $MetaTags = $HeadTag.childNodes| Where-Object {$_.nodename -eq 'META'}

    write-debug "Working on `"$($resultTable.title)`""

    # TODO: Get the MP description
    $resultTable.description = $metaTags  | Where-Object {$_.name -eq 'description'} | Select-Object -ExpandProperty content


    # Get the MP Category
    $Category = ($item.href -split "=")[1]

    # TODO: Get Monitoring Rule Element
    $Links = Get-ElementLinks -ElementType Rule


    $counter = 0
    $outputfolder = "C:\apps\skype"
    foreach ($link in $Links | ?{$_.href -match "Rule"} | ?{$_.href -notmatch "perf"}) {

        # unescape amperstand from link and follow it
        $href = $link.href -replace "&amp;","&"
        write-host "trying $($href)"

        $content = @()
        try { 
            $content = Invoke-WebRequest -Uri $href
        } catch { write-host "there was a problem with URI $($href)." ; return 1 }

        # get the content of the knolwedge base for the element
        $content_kb = $content.AllElements | ?{$_.class -match "ScrollArea KnowledgeBase"} | Select -ExpandProperty OuterText

        # get the content of the code for the element
        $content_code = $content.ParsedHtml.getElementsByTagName('PRE') | select -expand innertext
        $content_code = [xml]$content_code

        # extract useful elements of the code and place in object

        # process event log elements
        if ($content_code.Rule.Category -eq "EventCollection") {

            <#
            $counter++
            $outputfile = "$($outputfolder)\$($content_code.Rule.id).xml"
            Set-Content -Path $outputfile -Value $content_code.Rule.DataSources.DataSource.OuterXml

            if ($counter -eq 10) { break }
            #>

            $info = @{
                'MPName' = $resultTable.title | select
                'RuleName' = $link.outerText | select
                'id' = $content_code.rule.ID | select
                'Target' = $content_code.rule.Target | select
                'Category' = $content_code.Rule.Category | select
                'LogName' = $content_code.Rule.DataSources.DataSource.LogName | select
                'SourceName' = ($content_code.rule.DataSources.DataSource.Expression.and.Expression | ?{$_.SimpleExpression.ValueExpression.xpathquery.'#text' -eq "PublisherName"}).SimpleExpression.ValueExpression.Value | select-object -ExpandProperty InnerText
                'EventDisplayNumber' = ($content_code.rule.DataSources.DataSource.Expression.and.Expression | ?{$_.SimpleExpression.ValueExpression.xpathquery.'#text' -eq "EventDisplayNumber"}).SimpleExpression.ValueExpression.Value | select-object -ExpandProperty InnerText
                'KBArticle' = ($content_kb) | Select
                
            }
            if ($info.LogName -ne $null) { 
                $WinEventLogObjects += New-Object -TypeName PSObject -Property $info              
            }

        }


        # process performance monitor elements
        if ($content_code.Rule.Category -eq "PerformanceCollection") {
            $info = @{
                'MPName' = $resultTable.title | select
                'RuleName' = $link.outerText | select
                'id' = $content_code.rule.ID | select
                'Target' = $content_code.rule.Target | select
                'Category' = $content_code.Rule.Category | select
                'ObjectName' = $content_code.Rule.DataSources.DataSource.ObjectName | select
                'CounterName' = $content_code.Rule.DataSources.DataSource.CounterName | select
                'Frequency' = $content_code.Rule.DataSources.DataSource.Frequency | select
                'Tolerance' = $content_code.Rule.DataSources.DataSource.Tolerance | select
                'MaximumSampleSeparation' = $content_code.Rule.DataSources.DataSource.MaximumSampleSeparation | select
                'KBArticle' = $content_kb | select
            }
            if ($info.ObjectName -ne $null) { 
                $PerfmonObjects += New-Object -TypeName PSObject -Property $info        
            }
        }

    }

}


# Define the base output folder
$Outputfolder = "c:\apps\Skype"
if (!(Test-Path -Path $outputfolder)) {
    write-host "Output folder path not found: $($outputfolder)."
    exit
}

# get the unique components of Skype services.  There might should be a splunk app targeted to each component host
$Components = $WinEventLogObjects | Select-Object -Unique -Property Target
foreach ($Component in $Components.target) {

    # make sure the app folder exists
    $ComponentShort = ($Component -split "\.")[-1]
    #$appfolder = "$($outputfolder)\TA-SFB-$($ComponentShort)"  
    $appfolder = "$($outputfolder)\TA-SFB2019"  

    if (!(Test-Path -Path $appfolder)) { New-Item -Path $appfolder -ItemType Directory | Out-Null }   

    # make sure the app LOCAL folder exists
    $appFolderLocal = "$($appfolder)\local"
    if (!(test-path -path $appFolderLocal)) { New-Item -Path $appFolderLocal -ItemType Directory | Out-Null}

    # define path to inputs.conf
    $appInputsConf = "$($appFolderLocal)\inputs.conf"

    # get the eventlogs for this target
    $ComponentEvents = $WinEventLogObjects | ?{$_.target -eq $Component}

    # handle each unique log name
    foreach ($LogName in $ComponentEvents | Select-Object -Unique -ExpandProperty LogName) {

        $ComponentEventsForLogName = $ComponentEvents | ?{$_.LogName -eq $LogName}

        # Combine information into splunk inputs.conf stanza
        $Stanza = @()
        $Stanza += ""
        $Stanza += "# Event Collection Stanza for Skype $($ComponentShort) component."
        $Stanza += "[WinEventLog://$($LogName)]"
        $Stanza += "index = main"

        # Build whitelist of EventCodes to collect for each SourceName
        $ComponentSourceNames = $ComponentEventsForLogName | Select-Object -Unique SourceName
        $counter = -1
        foreach ($ComponentSourceName in $ComponentSourceNames.SourceName) {
            $counter++
            $ComponentSourceNameEvents = $ComponentEventsForLogName | ?{$_.SourceName -eq $ComponentSourceName}
            $EventCodes = $ComponentSourceNameEvents.EventDisplayNumber -join "|"
            $Stanza += "whitelist.$($counter) = SourceName=%^`($($ComponentSourceName)`)$% EventCode=%^`($($EventCodes)$`)%"
        }

        write-host "Writing $($Stanza.count) eventlog stanza lines to $($appInputsConf)."
        Add-Content -Path $appInputsConf -value $Stanza

    }

}


# get the perfmon logs of interest for this target
$Components = $PerfmonObjects | Select-Object -Unique -Property Target
foreach ($Component in $Components.target) {

    # make sure the app folder exists
    $ComponentShort = ($Component -split "\.")[-1]
    #$appfolder = "$($outputfolder)\TA-SFB-$($ComponentShort)"  
    $appfolder = "$($outputfolder)\TA-SFB2019"  
    if (!(Test-Path -Path $appfolder)) { New-Item -Path $appfolder -ItemType Directory  | Out-Null }

    # make sure the app LOCAL folder exists
    $appFolderLocal = "$($appfolder)\local"
    if (!(test-path -path $appFolderLocal)) { New-Item -Path $appFolderLocal -ItemType Directory | Out-Null}

    # define path to inputs.conf
    $appInputsConf = "$($appFolderLocal)\inputs.conf"

    # get the perfmon objects for this target
    $TargetPerfmonObjects = $PerfmonObjects | ?{$_.target -eq $Component}   

    # handle each unique log name
    foreach ($TargetPerfmonObject in $TargetPerfmonObjects | Select-Object -Unique -ExpandProperty ObjectName) {

        $TargetPerfmonObjectsForObjectName = $TargetPerfmonObjects | ?{$_.ObjectName -eq $TargetPerfmonObject}

        # Combine information into splunk inputs.conf stanza
        $Stanza = @()
        $Stanza += ""
        $Stanza += "# Event Collection Stanza for Skype $($ComponentShort) component."
        $Stanza += "[perfmon://$($TargetPerfmonObject)]"
        $Stanza += "index = main"
        $Stanza += "object = $($TargetPerfmonObject)"
        $Stanza += "instance = *"
      
        [string]$counters = $TargetPerfmonObjectsForObjectName.CounterName -join ";"
        $Stanza += "counter = $($counters)"

        write-host "Writing $($Stanza.count) perfmon stanza lines to $($appInputsConf)."
        Add-Content -Path $appInputsConf -value $Stanza

    }
}
