param(
$EmailTo,
$From_Email,
$Subject,
$cc,
$bcc
)
try
{



$Body = @"
<html>
<body style="font-family:Times New Roman">
<br/><img src='cid:head.png'<br/>
<br>
<br>
<br /><font face='Times New Roman'><b><i>Hi All, </a></i></b></font><br/>

<br /><font face='Times New Roman'><b><i>Please find the Disk space details </a></i></b></font><br/>

<br/><img src='cid:new.png'<br/>

<br /><font face='Times New Roman'><b><i>Regards </a></i></b></font><br/>
<br /><font face='Times New Roman'><b><i>3M Automation Center Team </a></i></b></font><br/>
<br>
<br>
<br/><img src='cid:footer.png'<br/>
</body>
</html>
"@



Send-MailMessage -SmtpServer "mailserv.mmm.com" -To $EmailTo -From $From_Email -Subject $Subject -BodyAsHtml -body $Body -Attachments "\\acdev01\3M_CAC\IPM_FSM\Mail_elements\head.png","\\acdev01\3M_CAC\IPM_FSM\Mail_elements\new.png", "\\acdev01\3M_CAC\IPM_FSM\Mail_elements\Footer.png" -Cc $cc -Bcc $bcc 



Write-Output "Mail sent to $EmailTo Successfully"
}
catch
{

Write-Output $_.Exception.Message



}