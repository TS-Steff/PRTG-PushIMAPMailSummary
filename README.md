# PRTG-PushIMAPMailSummary
*This project is as it is*

Summarizes mailcount by Subject with Error, Warning and Info
This is specific for QNAP HyperBackup-Mails. But could be easily modified for other mail subjects.

## Description
The script lookus up all mails in a Folder by IMAP and summarizes Mails with [Error], [Warn] or [Info] in the subject and sends the results to a PRTG-Push Sensor Advanced

It also checks the last received mail subject if it contains Error, Warn or Info. There is a PRTG Custom Lookup for this too.

The script uses MimeKit (http://www.mimekit.net/) and MailKit (https://github.com/jstedfast/MailKit)