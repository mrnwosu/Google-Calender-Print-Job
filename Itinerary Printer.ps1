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
        $isAllDay = $false
        $time = ""

        if($null -ne $_.start.date){
            $isAllDay = $true
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
    Write-Host "Starting IE"
    $ie = new-object -com "InternetExplorer.Application"


    Write-Host "Navigating to $fileName"
    $ie.Navigate($fileName)
    while ( $ie.busy ) { Start-Sleep -Seconds 3 }

    Write-Host "Executing Print"
    $ie.ExecWB(6,2)
    while ( $ie.busy ) { Start-Sleep -Seconds 3 }

    Write-Host "Waiting for print to send "
    Start-Sleep -Seconds 5
    $ie.quit()
}

function print-google-calendar-itinerary($targetEmail, $certPath, $certPassword, [string]$serviceAccount){
    if($null -eq (Get-Module -Name UMN-Google)){
        Write-Host "Installing Module UMN-Google"
        Install-Module UMN-Google -Force
    }

    Write-Host "Target Email Is $targetEmail"
    $email = $targetEmail

    Write-Host "Authenticating"
    $token = auth `
        -certPath $certPath `
        -certPassword $certPassword `
        -serviceAccount $serviceAccount.Trim()

    Write-Host "Authenticated"

    Write-Host "Getting Events"
    $eventResult = getTodaysEvents -token $token -email $email
    Write-Host "Events Retrieved"

    Write-Host "Generating HTML"
    $generatedHtml = generateHtml -eventResult $eventResult
    Write-Host "HTML Generated "

    $todaysDateString = Get-Date -Format "MMMM-dd-yyyy"
    $fileName = "$todaysDateString.html"
    $fullFileName = Join-Path (pwd) $fileName

    Write-Host "Itinerary Report is here: $fullFileName"
    $generatedHtml | Out-File $fullFileName

    Write-Host "Printing to Default printer with Internet Explorer..."
    printHtmlViaIE -fileName $fullFileName
    Write-Host "Printed"
}