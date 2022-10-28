Import-Module  "./Itinerary Printer.ps1";

# Check Environment Variable for "GoogleCertPasswood" if not use plain text
$certPassword = If ( $null -ne $env:GoogleCertPasswood ) { $env:GoogleCertPasswood } Else { "notasecret"}  


print-google-calendar-itinerary -targetEmail "< EMAIL ADDRESS >" `
    -certPath "< CERT PATH >" `
    -certPassword $certPassword `
    -serviceAccount "< SERVICE ACCOUNT >" `
