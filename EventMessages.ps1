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
    $search = "source=`"*WinEventLog:$($SelectedEvent.LogName)`" `"<EventID>$($SelectedEvent.EventID)</EventID>`" OR EventCode=`"$($SelectedEvent.EventID)`""  #primary search.. Can't assume eventcode is extracted yet.

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

# add type allowing interaction with eventlog data
Add-Type -AssemblyName System.Core

# Get the EventLogSession object
$EventSession = [System.Diagnostics.Eventing.Reader.EventLogSession]::GlobalSession;

$EventProviderNames = $EventSession.GetProviderNames()

# link providers to log names
if (!($LogNamesDB)) {
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

                    # skip analytic, debug and classic logs
                    if (($EventLogConfiguration.LogType -notmatch "(Analytical|Debug)") -and ($EventLogConfiguration.IsClassicLog -eq $false)) {

        
                        $info = @{
                            "LogName" = $LogLink.LogName
                            "ProviderName" = $EventProviderName
                            "IsEnabled" = $EventLogConfiguration.IsEnabled
                            "IsClassic" = $EventLogConfiguration.IsClassicLog
                            "Type" = $EventLogConfiguration.LogType                                          
                        }

                        $LogNamesDB += New-Object -TypeName PSObject -Property $info
                    }
                } catch {
                    Write-debug "Error accessing System.Diagnostics.Eventing.Reader.EventLogConfiguration properties of $($LogLink.LogName)."
                }

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
#        $metaData.LogLinks | ?{$_.LogName -match $Selection.LogName}

            # now make sure this link matches one the user selected.
#            if ($LogLink.LogName -eq $Selection.LogName) {

#                write-debug "Found link to selected [$($selection.LogName)] log sourced from [$($selection.ProviderName)] provider."

        $EventLogConfiguration = New-Object System.Diagnostics.Eventing.Reader.EventLogConfiguration -ArgumentList $LogLink.LogName

        foreach ($eventData in $metaData.Events) {

            $LogName = [string]($eventdata.LogLink).LogName

            if ($LogName -eq $Selection.LogName) { 

                # only append if there is a useful description
                if (($eventdata.Description) -and ($eventdata.Level.DisplayName -ne $null)) {

                    $info = @{
                        "Provider" =  $Selection.ProviderName

                        "LogName" = $LogName
                        "EventID" = $eventdata.Id
                        "Version" = $eventdata.Version
                        "Level" = $eventdata.Level.DisplayName
                        "Description" = $eventdata.Description -replace "\n","; "
                                
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


$EventDB = ($EventDB | Sort-Object Index | Get-Unique -AsString)

$SelectedEvents = $EventDB | Select-Object -Property LogName, Provider, Type, EventID, Version, Level, Description | sort-object -property Provider, LogName, EventID | Out-GridView -PassThru -Title "Select an EventLog ID of Interest" 

if (!($SelectedEvents)) { exit }

$PossibleActions = @("Export selected items to CSV file","Convert selected to Splunk inputs","Convert selected to Splunk searches")
$SelectedActions = $PossibleActions | Out-GridView -PassThru -Title "Select action(s) to take on selected items."


# get a random string for output filenames to share
$rando = get-random -Minimum 100 -Maximum 999

foreach ($SelectedAction in $SelectedActions) {

    if ($SelectedAction -match "CSV") {
        # do the CSV file task

        # prepare the temp file
        $EventMessagesOutputCSV = "$($env:temp)\EventMessages_csv_$($rando).csv"
        if (Test-Path -Path $EventMessagesOutputCSV) { Remove-Item -Path $EventMessagesOutputCSV -Force }

        # commit the output and display for user
        $SelectedEvents | Export-Csv -Path $EventMessagesOutputCSV -NoTypeInformation
        write-host "CSV file written to $($EventMessagesOutputCSV)."
        Start-Process -FilePath "Notepad.exe" -ArgumentList $EventMessagesOutputCSV
    }

    if ($SelectedAction -match "inputs") {
        # do the Splunk inputs file task

        # prepare the temp file
        $EventMessagesOutputInputs = "$($env:temp)\EventMessages_csv_$($rando).csv"
        if (Test-Path -Path $EventMessagesOutputInputs) { Remove-Item -Path $EventMessagesOutputInputs-Force }

        # prepare the content
        $Content = @()

        # group the selected events by sourcetype
        $LogNames = ($SelectedEvents | group LogName).Name

        foreach ($LogName in $LogNames) {
            $Events = $SelectedEvents | ?{$_.LogName -eq "$($LogName)"}
            $Content += ""
            $Content += "[WinEventLog://$($LogName)]"
            $Content += "index = main"
            
            # build a list of EventTypes to include in a sample filtering array
            $NameTypes="" 
            $Events | group Level | select -expandproperty Name | %{if ($NameTypes -eq "") { $NameTypes += $_ } else { $NameTypes += "|" + $_ }} ; $NameTypes = "(" + $NameTypes + ")"
            $content += "#whitelist.1 = Type=%^$($NameTypes)$%"

            # build a list of EventCodes to include in a sample filtering array
            $eEventCodes = ""
            $Events | group EventID | Select -ExpandProperty Name | %{if ($eEventCodes -eq "") { $eEventCodes += $_ } else { $eEventCodes += "|" + $_ }} ; $eEventCodes = "(" + $eEventCodes + ")" 
            $content += "#blacklist.1 = EventCode=%^$($eEventCodes)$%"

            # build a list of event descriptions to reference in inputs file
            $eDescriptions = @()
            foreach ($Event in $Events) {
                $eDescriptions += "# EventID: $($Event.EventID), Level: $($Event.Level), Description: $($Event.Description)"
            }
            $content += $eDescriptions
            Add-Content -Path $EventMessagesOutputInputs -value $Content
        }
      
        # commit the output and display for user
        write-host "Splunk inputs written to $($EventMessagesOutputInputs)."
        Start-Process -FilePath "Notepad.exe" -ArgumentList $EventMessagesOutputInputs

    }

    if ($SelectedAction -match "searches") {
        # do the Splunk searches 

        # prepare the temp file
        $EventMessagesOutputSearch = "$($env:temp)\EventMessages_searches_$($rando).csv"
        if (Test-Path -Path $EventMessagesOutputSearch) { Remove-Item -Path $EventMessagesOutputSearch -Force }

        # prepare the content
        $Content = @()
        foreach ($SelectedEvent in $SelectedEvents) {
            $Search = ConvertTo-SplunkSearch -SelectedEvent $SelectedEvent
            $Content += "`n$($search)"
        }
        
        # commit the output and display for user
        Add-Content -Path $EventMessagesOutputSearch -value $Content
        write-host "Splunk SPL statements written to $($EventMessagesOutputSearch)."
        Start-Process -FilePath "Notepad.exe" -ArgumentList $EventMessagesOutputSearch
    }
}


#>
