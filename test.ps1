Import-Module  "./Itinerary Printer.ps1";

print-google-calendar-itinerary -targetEmail "ike.nwosu.151@gmail.com" `
    -certPath ".\get-calenday-events-8bca31a28b28.p12" `
    -certPassword "notasecret" `
    -serviceAccount "ps-cal-list@get-calenday-events.iam.gserviceaccount.com" `
    -printerName "$ Test Printer"