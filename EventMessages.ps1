function ConvertTo-SplunkSearch {
    param($SelectedEvent)
    # Transform event into splunk search:
    $search = "source=`"*WinEventLog:$($SelectedEvent.LogName)`" `"<EventID>$($SelectedEvent.Id)</EventID>`" OR EventCode=`"$($SelectedEvent.Id)`""  #primary search.. Can't assume eventcode is extracted yet.

    # extract the "data" fields that are unique to each event id but defined in provider template.
    $dataitems = $SelectedEvent.Template -split "`n" | Select-String -Pattern "data name="
    $fields = ""
    foreach ($dataitem in $dataitems) {
        $Data = ([regex]"<data name=`"([^`"]+)`"").match($dataitem).groups[1].value
        if ($fields -eq "") {
            $fields = "$($Data)"
        } else {
            $fields += ", $($Data)"
        }
    }
    $search += "`n | table _time host source EventID $($fields) `n | sort 0 - _time"
    return $search
}

function ConvertTo-SplunkView {
    param($SelectedEvent,$search,$label)
    $view = "<form theme=`"light`">"
    $view += "`n  <label>$($label)</label>"
    $view += "`n  <fieldset submitButton=`"true`">"
    $view += "`n    <input type=`"time`" token=`"field1`">"
    $view += "`n      <label></label>"
    $view += "`n      <default>"
    $view += "`n        <earliest>-7d@h</earliest>"
    $view += "`n        <latest>now</latest>"
    $view += "`n      </default>"
    $view += "`n    </input>"
    $view += "`n  </fieldset>"
    $view += "`n  <row>"
    $view += "`n    <panel>"
    $view += "`n      <table>"
    $view += "`n        <search>"
    $view += "`n          <query>"
    $view += "`n          $($search)"
    $view += "`n		  </query>"
    $view += "`n          <earliest>`$field1.earliest$</earliest>"
    $view += "`n          <latest>`$field1.latest$</latest>"
    $view += "`n          <sampleRatio>1</sampleRatio>"
    $view += "`n        </search>"
    $view += "`n        <option name=`"count`">20</option>"
    $view += "`n        <option name=`"dataOverlayMode`">none</option>"
    $view += "`n        <option name=`"drilldown`">none</option>"
    $view += "`n        <option name=`"percentagesRow`">false</option>"
    $view += "`n        <option name=`"refresh.display`">progressbar</option>"
    $view += "`n        <option name=`"rowNumbers`">true</option>"
    $view += "`n        <option name=`"totalsRow`">false</option>"
    $view += "`n        <option name=`"wrap`">true</option>"
    $view += "`n      </table>"
    $view += "`n    </panel>"
    $view += "`n  </row>"
    $view += "`n</form>"
    return $view
}

$output_path = "$($env:temp)\dashboards"
if (!(Test-Path -Path $output_path)) { mkdir -Path $output_path -ErrorAction SilentlyContinue }

# Get the EventLogSession object
$EventSession = [System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession;

# Get the ETW provider names
$EventProviderNames = $EventSession.GetProviderNames()
$ProviderMetadataList = @()
$MenuItems = @()
foreach ($EventProviderName in $EventProviderNames) {
    Try {
        $ProviderMetadataList += New-Object -TypeName System.Diagnostics.Eventing.Reader.ProviderMetadata -ArgumentList $EventProviderName -ErrorAction SilentlyContinue
        $MenuItem = @()
        foreach ($Link in $ProviderMetadataList[-1].LogLinks) {
            $CustomEvent = new-object -TypeName PSObject
            $CustomEvent | Add-Member -MemberType NoteProperty -Name LogName -Value $Link.LogName
            $CustomEvent | Add-Member -MemberType NoteProperty -Name Provider -Value $ProviderMetadataList[-1].Name
            $MenuItem += $CustomEvent
        }
        $MenuItems += $MenuItem

    } catch {
        Write-Debug "Problem getting provider metadata for $($EventProviderName)."
    }
}

# Prompt user to select one to explore
$SelectedLogName  = $MenuItems | Select-Object -Unique -Property LogName | Sort-Object -Property LogName  | Out-GridView -PassThru -Title "Select an EventLog Provider of Interest"

# Get associated providers for selected log
$AssociatedProviders = $Menuitems | ?{$_.LogName -eq $SelectedLogName.LogName}

# Get metadata for selected providers
$Events  = @()
foreach ($ProviderMetadataItem in $ProviderMetadataList) {
    foreach ($ProviderMetadataItemLogLink in $ProviderMetadataItem.LogLinks) {
        foreach ($AssociatedProvider in $AssociatedProviders) {
            if ($ProviderMetadataItemLogLink.LogName -match $AssociatedProvider.LogName) {
                foreach ($Event in $ProviderMetadataItem.Events) {

                    $Event.Events | Select-Object -Property Id, OpCode, Task, Keywords, Description, Template, LogLink

                    $CustomEvent = new-object -TypeName PSObject
                    $CustomEvent | Add-Member -MemberType NoteProperty -Name "LogName" -Value "$($AssociatedProvider.LogName)"
                    $CustomEvent | Add-Member -MemberType NoteProperty -Name "Provider" -Value "$($AssociatedProvider.Provider)"
                    $CustomEvent | Add-Member -MemberType NoteProperty -Name "Id" -Value "$($Event.Id)"
                    $CustomEvent | Add-Member -MemberType NoteProperty -Name "OpCode" -Value "$($Event.Opcode.DisplayName)"
                    $CustomEvent | Add-Member -MemberType NoteProperty -Name "Task" -Value "$($Event.Task.DisplayName)"
                    $CustomEvent | Add-Member -MemberType NoteProperty -Name "Keywords" -Value "$($Event.Keywords.DisplayName)"
                    $CustomEvent | Add-Member -MemberType NoteProperty -Name "Description" -Value "$($Event.Description)"
                    $CustomEvent | Add-Member -MemberType NoteProperty -Name "Template" -Value "$($Event.Template)"

                    $Events += $CustomEvent

                }

            }
       }
    }
}

# Prompt user to select one to explore
$SelectedEvents = $Events | Sort-Object -Property Id | Out-GridView -PassThru -Title "Select an EventLog ID of Interest" 

foreach ($SelectedEvent in $SelectedEvents) {
    $Description = (($SelectedEvent.Description) -split "`n")[0]
    $Description = (($Description) -split ":")[0]
    $Description = (($Description) -split "\.")[0].trim()

    $label = "WinEventLog:$($SelectedEvent.LogName):$($SelectedEvent.Id):$($Description)" 

    $Search = ConvertTo-SplunkSearch -SelectedEvent $SelectedEvent
    $View = ConvertTo-SplunkView -SelectedEvent $SelectedEvent -search ([Security.SecurityElement]::Escape($Search)) -label $label


    # remove illegal chars for filenames
    $labelEscape = $label -replace "\\","_"
    $labelEscape = $labelEscape -replace "/","_"
    $labelEscape = $labelEscape -replace ":","_"
    $labelEscape = $labelEscape -replace "\*","_"
    $labelEscape = $labelEscape -replace "\?","_"
    $labelEscape = $labelEscape -replace "\`"","_"
    $labelEscape = $labelEscape -replace "`'","_"
    $labelEscape = $labelEscape -replace "<","_"
    $labelEscape = $labelEscape -replace ">","_"
    $labelEscape = $labelEscape -replace "\|","_"
    $labelEscape = $labelEscape -replace "\%","_"


    $output_filename = $labelEscape
    $output_filepath = "$($output_path)\$($output_filename).xml"
    if (Test-Path -Path $output_filepath) { Remove-Item -Path $output_filepath -Force }

    write-host "Writing exploratory Splunk dashboard for [$($SelectedEvent.LogName)] EventID [$($SelectedEvent.Id)] to [$($output_filepath)]."
    Add-Content -Path $output_filepath -value $View
}
