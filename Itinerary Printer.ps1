function auth($certPath, $certPassword, $serviceAccount){
    $scope = "https://www.googleapis.com/auth/calendar.readonly"
    try {
        $accessToken = Get-GOAuthTokenService -scope $scope -certPath $certPath -certPswd $certPassword -iss $serviceAccount
        return $accessToken
    } 
    catch {
        $err = $item_.Exception
        $err | Select-Object -Property *
        "Response: "
        $err.Response
        throw $err
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

function print-google-calendar-itinerary($targetEmail, $certPath, $certPassword, $serviceAccount, $printerName){
    if($null -eq (Get-Module -Name UMN-Google)){
        Install-Module UMN-Google
    }

    $email = $targetEmail
    $token = auth `
        -certPath $certPath `
        -certPassword $certPassword `
        -serviceAccount $serviceAccount

    $eventResult = getTodaysEvents -token $token -email $email
    $generatedHtml = generateHtml -eventResult $eventResult 
    $todaysDateString = Get-Date -Format "MMMM-dd-yyyy"
    $fileName = "$todaysDateString.html"
    $fullFileName = Join-Path (pwd) $fileName
    $generatedHtml | Out-File $fullFileName
    Get-Content -Path $fullFileName | Out-Printer -Name $printerName
}