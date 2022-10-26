Import-Module  "./Itinerary Printer.ps1";

print-google-calendar-itinerary -targetEmail "< EMAIL ADDRESS >" `
    -certPath "< CERT PATH >" `
    -certPassword "< CERT SECRET >" `
    -serviceAccount "< SERVICE ACCOUNT >" `
    -printerName "< PRINTER NAME >"