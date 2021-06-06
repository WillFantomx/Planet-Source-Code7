<div align="center">

## Error Handling VB6 \( Similar to Windows XP Error Message\)


</div>

### Description

Demo application developed similar to error screen shown in Windows XP

This module will be used in error handling for VB application.

This application will log an error log and send an email notification.

Modules to be included in client project :

1. FrmErrorMsg.frm

2. moduleGlobal.bas

Working of errormsg :

If each method add below code..

On Error GoTo ErrorHandler

--

--

--

--

Exit Sub/function

ErrorHandler:

Call frmErrorMsg.SetErrorLog(Err.Number, Err.Description, "Command1_Click", "frmDemo")

frmErrorMsg.Show vbModal

End

Where command1_Click is method name and frmdemo is form name.

So a log file is created with this log and a custom form is shown.

Please modify form as per requirement.

If user clicks on send Error Report. Error log is send to default emailaddress. So change email address first before sending mail.

Also you can either automatically send email without knowing user about it. Just pass first parameter to method SendMessage as "false"

Another alternative to this is sending data on site. for e.g. www.yoursite.com/submit.aspx?... etc.
 
### More Info
 
This module will be used in error handling for VB application.

This application will log an error log and send an email notification.

Modules to be included in client project :

1. FrmErrorMsg.frm

2. moduleGlobal.bas

Working of errormsg :

If each method add below code..

On Error GoTo ErrorHandler

--

--

--

--

Exit Sub/function

ErrorHandler:

Call frmErrorMsg.SetErrorLog(Err.Number, Err.Description, "Command1_Click", "frmDemo")

frmErrorMsg.Show vbModal

End

Where command1_Click is method name and frmdemo is form name.

So a log file is created with this log and a custom form is shown.

Please modify form as per requirement.

If user clicks on send Error Report. Error log is send to default emailaddress. So change email address first before sending mail.

Also you can either automatically send email without knowing user about it. Just pass first parameter to method SendMessage as "false"

Another alternative to this is sending data on site. for e.g. www.yoursite.com/submit.aspx?... etc.


<span>             |<span>
---                |---
**Submitted On**   |2005-08-30 15:38:46
**By**             |[Amod](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/amod.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Error\_Hand192887912005\.zip](https://github.com/Planet-Source-Code/amod-error-handling-vb6-similar-to-windows-xp-error-message__1-62385/archive/master.zip)








