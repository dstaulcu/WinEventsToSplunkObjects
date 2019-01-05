<# TODO:
if logtype is classic, get localized $metaData.MessageFilePath
e.g. c:\windows\syswow64\en-us\esent.dll.mui 
The file has a message table which contains the ID (in hex) and Messages (as strings).  
Unforunately, ID/message levels are not defined for classic logs

https://nsis.sourceforge.io/Reading_Message_Table_resources_from_DLL_files
GetSystemDefaultLangID --> LoadLibraryExW --> FormatMessageW --> FreeLibrary
#>


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

# add type allowing interaction with eventlog data
Add-Type -AssemblyName System.Core

# Get the EventLogSession object
$EventSession = [System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession;

$EventProviderNames = $EventSession.GetProviderNames()

# link providers to log names
$LogNamesDB = @()
$Progress = 0
foreach ($EventProviderName in $EventProviderNames) {

    $Progress++

    Write-Progress -activity "Building list log names linked from log providers." -status "Evaluating provider $($Progress) of $($EventProviderNames.count) - [$($EventProviderName)]." -PercentComplete $(($Progress/$EventProviderNames.count)*100)
    
    Try {
        $metaData = New-Object -TypeName System.Diagnostics.Eventing.Reader.ProviderMetadata -ArgumentList $EventProviderName
    } catch {
        Write-debug "Error accessing Eventing.Reader.ProviderMetadata properties of $($EventProviderName)."
        continue
    }

    # if there is not metadata or log link, skip
    if (($metaData) -and ($metadata.LogLinks)) {

        # some providers are linked to multiuple log types, check each link
        foreach ($LogLink in $metaData.LogLinks) { 

            Try {
                $EventLogConfiguration = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration -ArgumentList $LogLink.LogName
        
                $info = @{
                    "LogName" = $LogLink.LogName
                    "ProviderName" = $EventProviderName
                    "IsEnabled" = $EventLogConfiguration.IsEnabled
                    "IsClassic" = $EventLogConfiguration.IsClassicLog
                    "Type" = $EventLogConfiguration.LogType                                          
                }

                $LogNamesDB += New-Object -TypeName PSObject -Property $info
            } catch {
                Write-debug "Error accessing System.Diagnostics.Eventing.Reader.EventLogConfiguration properties of $($LogLink.LogName)."
            }

        }
    }
}

# Present the user a list of distinct logs they want to examine event details of.
#?{$_.IsClassic -eq $True} | 
[array]$Selections = $LogNamesDB | Select-object LogName, ProviderName, Type, IsEnabled, IsClassic | sort-object -property LogName, ProviderNane | Out-GridView -Title "Select one or more log files to extract event id details from." -PassThru

# Now build a database of the event id details of each log selected
$EventDB = @()
$Progress = 0
foreach ($Selection in $Selections) {
    $Progress++

    Write-Progress -activity "Loading event log metadata from provider of selected logs." -status "Evaluating log $($Progress) of $($Selections.count) - [$($Selection)]." -PercentComplete $(($Progress/$Selections.count)*100)

    write-debug "Looking for [$($selection.LogName)] log sourced from [$($selection.ProviderName)] provider."

    Try {
        $metaData = New-Object -TypeName System.Diagnostics.Eventing.Reader.ProviderMetadata -ArgumentList $Selection.ProviderName
    } catch {
        Write-debug "Error accessing Eventing.Reader.ProviderMetadata properties of $($Provider)."
        continue
    }

    # if there is not metadata or log link, skip
    if (($metaData) -and ($metadata.LogLinks)) {

        # some providers are linked to multiuple log types, check each link
        foreach ($LogLink in $metaData.LogLinks) {       

            # now make sure this link matches one the user selected.
            if ($LogLink.LogName -eq $Selection.LogName) {

                write-debug "Found link to selected [$($selection.LogName)] log sourced from [$($selection.ProviderName)] provider."

                $EventLogConfiguration = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration -ArgumentList $LogLink.LogName

                foreach ($eventData in $metaData.Events) {

                    # only append if there is a useful description
                    if (($eventdata.Description) -and ($eventdata.Level.DisplayName -ne $null)) {

                        $info = @{
                            "Provider" =  $Selection.ProviderName

                            "LogName" = $LogLink.LogName
                            "EventID" = $eventdata.Id
                            "Version" = $eventdata.Version
                            "Level" = $eventdata.Level.DisplayName
                            "Description" = $eventdata.Description
                                
                            "IsEnabled" = $EventLogConfiguration.IsEnabled
                            "IsClassic" = $EventLogConfiguration.IsClassicLog                               
                            "Type" = $EventLogConfiguration.LogType

                            "Index" = "$($LogLink.LogName) $($Selection.ProviderName) $($EventLogConfiguration.LogType) $($eventdata.Id) $($eventdata.Version)"

                        }

                        $EventDB += New-Object -TypeName PSObject -Property $info

                    }

                }
            }
        }
    }
}


$EventDB = ($EventDB | Sort-Object Index | Get-Unique -AsString)

$SelectedEvents = $EventDB | Sort-Object -Property Id | Select-Object -Property LogName, Provider, Type, EventID, Version, Level, Description | sort-object -property Provider, Type, LogName, EventID | Out-GridView -PassThru -Title "Select an EventLog ID of Interest" 

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

