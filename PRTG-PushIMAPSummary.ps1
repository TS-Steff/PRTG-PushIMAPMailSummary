<#
.NOTES
┌─────────────────────────────────────────────────────────────────────────────────────────────┐ 
│ ORIGIN STORY                                                                                │ 
├─────────────────────────────────────────────────────────────────────────────────────────────┤ 
│   DATE        : 2021.05.10                                                                  |
│   AUTHOR      : TS-Management GmbH, Stefan Müller                                           | 
│   DESCRIPTION : Push num files in IMAP Folder by subject with error, warning and info       |
|   HISTORY     : 2021.08.16 - Added Last message subject error, warn or info                 |
|                            - imap disconnect added                                          |
|                 2021.08.18 - Added Variable to summarize only unread messages               |
└─────────────────────────────────────────────────────────────────────────────────────────────┘
#>


#####
# Config
#####
$IMAP_HOST = ''
$IMAP_PORT = 993
$IMAP_USERNAME = ''
$IMAP_PASSWORD = ''
$IMAP_FOLDER = 'INBOX/Customer'
$DAYSToSummarize = 3
$UNSEEN_MAILS_ONLY = 1

$PRTG_PROBE = "" #include https or http
$PRTG_PORT = ""
$PRTG_KEY = ""


$PRTG_JSON = ""


if ($PSVersionTable.PSVersion.Major -lt 5){write-host "ERROR: Minimum Powershell Version 5.0 is required!" -F Yellow; return}  

function sendPush(){
    Add-Type -AssemblyName system.web

    write-host "result"-ForegroundColor Green
    write-host $PRTG_JSON 

    #$Answer = Invoke-WebRequest -Uri $NETXNUA -Method Post -Body $RequestBody -ContentType $ContentType -UseBasicParsing
    $answer = Invoke-WebRequest `
       -method POST `
       -URI ($PRTG_PROBE + ":" + $PRTG_PORT + "/" + $PRTG_KEY) `
       -ContentType "application/json" `
       -Body $PRTG_JSON `
       -usebasicparsing

    if ($answer.statuscode -ne 200) {
       write-warning "Request to PRTG failed"
       write-host "answer: " $answer.statuscode
       exit 1
    }
    else {
       $answer.content
    }
}


# Mailkit DLLs
# DLLs über http://www.mimekit.net/ oder über NuGet
"$PSScriptRoot\BouncyCastle.Crypto.dll","$PSScriptRoot\MimeKit.dll","$PSScriptRoot\MailKit.dll" | %{
    Unblock-File -Path $_
    Add-Type -Path $_ -EA Stop
}

# TLS Protokolle festlegen
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::GetNames([System.Net.SecurityProtocolType])

try{
    # IMAP Client
    $imap = New-Object MailKit.Net.Imap.ImapClient

    # verbindung
    $imap.Connect($IMAP_HOST,$IMAP_PORT,[MailKit.Security.SecureSocketOptions]::Auto)

    # Authentifizierung
    $imap.Authenticate($IMAP_USERNAME,$IMAP_PASSWORD)

    # Ordner öffnen
    $inbox = $imap.GetFolder($IMAP_FOLDER)
    $inbox.Open([MailKit.FolderAccess]::ReadWrite)

    # Alle Nachrichten im Ordner
    $msgLastXDays_ids = $inbox.Search([MailKit.Search.SearchQuery]::All)    
    write-host $msgLastXDays_ids.Count " Nachrichten"

    # Alle Nachrichten der letzten X Tage
    $dateFrom = (get-date).AddDays(-$DAYSToSummarize)
    $lastDaysMSG = $inbox.Search([MailKit.Search.SearchQuery]::DeliveredAfter($dateFrom))
    $numLastDaysMSG = $lastDaysMSG.Count
    #write-host $lastDaysMSG.Count " Nachreichten jünger als " $DAYSToSummarize " Tage"


    # Ungelesen Nachrichten der letzten X Tage
    $msgUnreadXDays_ids = $inbox.Search([MailKit.Search.SearchQuery]::NotSeen.And([MailKit.Search.SearchQuery]::DeliveredAfter($dateFrom)))
    #write-host $msgUnreadXDays_ids.Count " ungelesene Nachrichten in den letzten $DAYSToSummarize Tagen"

    # Nachrichten mit [Error] im Betreff der letzten X Tage
    if($UNSEEN_MAILS_ONLY = 1){
        $msgError_ids = $inbox.Search([MailKit.Search.SearchQuery]::SubjectContains("[Error]").And([MailKit.Search.SearchQuery]::DeliveredAfter($dateFrom)).And([MailKit.Search.SearchQuery]::NotSeen))
    }else{
        $msgError_ids = $inbox.Search([MailKit.Search.SearchQuery]::SubjectContains("[Error]").And([MailKit.Search.SearchQuery]::DeliveredAfter($dateFrom)))
    }
    $numErr = $msgError_ids.Count 
    #write-host $msgError_ids.Count " Fehlermeldungen in den letzten $DAYSToSummarize Tagen "

    # Nachrichten mit [Info] im Betreff der letzten X Tage
    if($UNSEEN_MAILS_ONLY = 1){
        $msgInfo_ids = $inbox.Search([MailKit.Search.SearchQuery]::SubjectContains("[Info]").And([MailKit.Search.SearchQuery]::DeliveredAfter($dateFrom)).And([MailKit.Search.SearchQuery]::NotSeen))
    }else{
        $msgInfo_ids = $inbox.Search([MailKit.Search.SearchQuery]::SubjectContains("[Info]").And([MailKit.Search.SearchQuery]::DeliveredAfter($dateFrom)))
    }
    $numInfo = $msgInfo_ids.Count 

    #write-host $msgInfo_ids.Count " Infonachrichten in den letzten $DAYSToSummarize Tagen "

    # Nachrichten mit [Error] im Betreff der letzten X Tage
    if($UNSEEN_MAILS_ONLY = 1){
        $msgWarn_ids = $inbox.Search([MailKit.Search.SearchQuery]::SubjectContains("[Warning]").And([MailKit.Search.SearchQuery]::DeliveredAfter($dateFrom)).And([MailKit.Search.SearchQuery]::NotSeen))
    }else{
        $msgWarn_ids = $inbox.Search([MailKit.Search.SearchQuery]::SubjectContains("[Warning]").And([MailKit.Search.SearchQuery]::DeliveredAfter($dateFrom)))
    }
    $numWarn = $msgWarn_ids.Count   

    #####
    # Betreff der letzten Nachricht auslesen
    #####
    $msgLastReceived = $inbox.GetMessage($msgLastXDays_ids.Count -1)

    if($msgLastReceived.Subject -Match "Error"){
        $lastMSGState = 2
    }elseif($msgLastReceived.Subject -Match "Warning"){ 
        $lastMSGState = 1 
    }elseif($msgLastReceived.Subject -Match "Info"){ 
        $lastMSGState = 0
    }else{
        $lastMSGState = "unknown"
    }


    $imap.Disconnect($true)

# Create PRTG_JSON

$PRTG_JSON = @"
{
    "prtg":{
        "result":[
            {
                "channel":"Days to lookup",
                "unit":"Custom",
                "value": $DAYSToSummarize
            },
            {
                "channel":"Messages in lookupdays",
                "unit":"Custom",
                "value": $numLastDaysMSG,
                "showChart":1,
                "showTable":1
            },
            {
                "channel":"Errors",
                "unit":"Custom",
                "value":$numErr,
                "showChart":1,
                "showTable":1,
                "LimitMaxError":1,
                "LimitErrorMsg": "$numErr errors in the last $DAYSToSummarize Days",
                "LimitMode":1

            },
            {
                "channel":"Warn Mails",
                "unit":"Custom",
                "value":$numWarn,
                "showChart":1,
                "showTable":1,
                "LimitMaxWarning":1,
                "LimitWarningMsg": "$numWarn warnings in the last $DAYSToSummarize Days",
                "LimitMode":1
            },
            {
                "channel":"Information Mails",
                "unit":"Custom",
                "value":$numInfo,
                "showChart":1,
                "showTable":1
            },
            {
                "channel":"Last message state",
                "unit":"Custom",
                "value":"$lastMSGState",
                "valueLookup":"ts.imap.qnap.mail",
                "showChart":1,
                "showTable":1
            }
        ]
    }
}
"@

    #write-host $PRTG_JSON
    sendPush

}catch{
    throw $_
}finally{
    # Disconnect
    if ($imap.Connected){
        $imap.Disconnect($true)
    }
}