function auth($certPath, $certPassword, $serviceAccount){
    $scope = "https://www.googleapis.com/auth/calendar.readonly"
    try {
        $accessToken = Get-GOAuthTokenService -scope $scope -certPath $certPath -certPswd $certPassword -iss $serviceAccount
        return $accessToken
    } 
    catch {
        throw $PSItem
    }
}

function getTodaysEvents($token, $email){
    $start = Get-Date -Hour "00" -Minute "00" -Second "00"
    $end = Get-Date -Hour "23" -Minute "59" -Second "59"

    $startTick = [Xml.XmlConvert]::ToString(($start),[Xml.XmlDateTimeSerializationMode]::Local)
    $endTick = [Xml.XmlConvert]::ToString(($end),[Xml.XmlDateTimeSerializationMode]::Local)

    $requestUri = "https://www.googleapis.com/calendar/v3/calendars/$email/events?timeMin=$startTick&timeMax=$endTick"
    return  Invoke-RestMethod -Headers @{"Authorization" = "Bearer $token"} -Uri $requestUri -Method Get -ContentType 'application/json' 
}


function getEventHtmlList($items){
    $list = New-Object Collections.Generic.List[string]
    $items | Sort-Object -Property @{Expression={$_.start.datetime}} | ForEach-Object {
        $html = "<div class=""event-wrapper"">"
        $time = ""

        if($null -ne $_.start.date){
            $time = "All Day"
        }
        else{
            $time = $_.start.datetime.ToString("hh:mm:tt") 
        }
        
        $description = $_.description
        $location = $_.location
        $summary = $_.summary
        $attendies = $_.attendees

        $html += "<h3>$time - $summary</h3>"
        if($null -ne $description){
            $html += "<p><b>Description</b>: $description</p>"
        }
        if($null -ne $location){
            $html += "<p><b>Location</b> :$location</p>"
        }
    
        if($null -ne $attendies -and $attendies.Length -gt 0){
            $html += "<p><b>Attendees</b>: "
            $a = ""
            $_.attendees | ForEach-Object {
                if($_.self -ne $true){
                    $person = $_.email
                    if($a.Length -gt 0){
                        $a += ", "
                    }
                    $a += $person
                }
            }
            $html += "$a</p>"
        }
        $html += "</div>"

        $list.Add($html)
    }
    return $list
}

function getNoItemsHtml{
    return "<h3>There are no events scheduled for today.</h3>"
}

function generateHtml($eventResult){
    $html = @"
    <html>
    <head>
        <style>
            body{
                font-family: Helvetica;
            }
            .wrapper{
                width: 80%;
                margin: auto;
            }
            .title{
                border-bottom: 10px;
                border-bottom-style: solid;
                border-radius: 3px;
                border-color: rgb(124, 124, 223);
            }
            .event-wrapper {
                margin: 20px 0px;
            }

            .event-wrapper h3 {
                margin: 0 4px 0 0
            }

            .event-wrapper p {
                margin: 4px 0px 
            }
        </style>
    </head>
    <body>
    <div class="wrapper">
"@

    $today = Get-Date -Format "MMMM, dd, yyyy"
    $html += "<h1 class=""title""> Itinerary for $today</h1>"
    if($null -eq $eventResult.items -or $eventResult.items.Count -eq 0){
        $html += getNoItemsHtml
    }
    else {
        getEventHtmlList($eventResult.items)  | ForEach-Object {
            $html += $_
        }

    }

    $html += "</div></body></html>"

    return $html
}

function printHtmlViaIE($fileName){
    log "Starting IE"
    $ie = new-object -com "InternetExplorer.Application"

    log "Navigating to $fileName"
    $ie.Navigate($fileName)
    while ( $ie.busy ) { Start-Sleep -Seconds 3 }

    log "Executing Print"
    $ie.ExecWB(6,2)
    while ( $ie.busy ) { Start-Sleep -Seconds 3 }

    log "Waiting for print to send "
    Start-Sleep -Seconds 5
    $ie.quit()
}

function print-google-calendar-itinerary($targetEmail, $certPath, $certPassword, [string]$serviceAccount){
    try{
        if($null -eq (Get-Module -Name UMN-Google)){
            log "Installing Module UMN-Google"
            Install-Module UMN-Google -Force
        }

        log "Target Email Is $targetEmail"
        $email = $targetEmail

        log "Authenticating"
        $token = auth `
            -certPath $certPath `
            -certPassword $certPassword `
            -serviceAccount $serviceAccount.Trim()

        log "Authenticated"

        log "Getting Events"
        $eventResult = getTodaysEvents -token $token -email $email
        log "Events Retrieved"
        $eventCount = $eventResult.items.Count
        log "There are $eventCount(s) scheduled for today."

        log "Generating HTML"
        $generatedHtml = generateHtml -eventResult $eventResult
        log "HTML Generated "

        $todaysDateString = Get-Date -Format "MMMM-dd-yyyy"
        $fileName = "$todaysDateString.html"
        $fullFileName = Join-Path (pwd) $fileName

        log "Itinerary Report is here: $fullFileName"
        $generatedHtml | Out-File $fullFileName

        log "Printing to Default printer with Internet Explorer..."
        printHtmlViaIE -fileName $fullFileName
        log "Printed"
    }
    catch{
        log "An error occurred:"
        log -message $_ -level "ERROR"
    }
    finally{
        $cwd = (pwd)
        log "Logging file is located at ""$cwd\job.log"""
    }
}

function log($message, $level = 'INFO'){
    $time = (Get-Date -Format "MM-dd-yy hh:mmtt")
    $log = "$time - $level - $message"
    
    if ($level -eq "ERROR") {
        Write-Error $log
    }
    else{
        Write-Host $log
    }    
    Add-Content job.log $log
}