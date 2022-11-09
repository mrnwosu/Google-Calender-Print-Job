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

$testHtml = @"
<html>
<body>
    <div style="position: relative; height: 150px; width: 300px; margin: auto; background-color: rgb(105, 105, 233, .3);">
        <div style="position: relative; margin: 50px auto; width: auto; text-align: center;  background-color: rgb(105, 105, 233, .3);">
            <h1 style="margin: auto; text-align: center">Test Page</h1>
        </div>
    </div>
</body>
</html>
"@

$testFileName = "textHtml.html" 

$testHtml >> $testFileName

printHtmlViaIE -fileName $testFileName

Remove-Item $testFileName