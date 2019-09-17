option explicit
Dim oEmail, oConf
Dim authuser, authpass, smtpserver
Dim recipient,sender,subject,message
Dim strComputer
Dim strUser
Dim strUserQualified
Dim NIC1, Nic, StrIP

strComputer = CreateObject("WScript.Network").ComputerName
strUser = CreateObject("WScript.Network").UserName
strUserQualified = CreateObject("WScript.Network").UserDomain & "\" & CreateObject("WScript.Network").UserName
Set NIC1 = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1
if Nic.IPEnabled then
StrIP = Nic.IPAddress(0)
end if
next
 
authuser="**ENTER_USER_HERE**"
authpass="**ENTER_PASSWORD_HERE**"
smtpserver="smtp.googlemail.com"
 
recipient="**ENTER_RECEIPT_MAIL_HERE"
sender="**ENTER_SENDER_MAIL_HERE**"
subject="Zugriff USB 1"
message="Computer: "& strComputer & vbNewLine &" User: "& strUser & vbNewLine &" Domainuser: "& strUserQualified & vbNewLine & " IP: " & strIP
 
Set oEmail = CreateObject("CDO.Message")
Set oConf = CreateObject("CDO.Configuration")
 
oConf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
oConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
oConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpserver
oConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
oConf.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = authuser
oConf.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = authpass
oConf.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
oConf.Fields.Update
 
oEmail.Configuration = oConf
oEmail.To = recipient
oEmail.From = sender
oEmail.ReplyTo = sender
oEmail.Subject = subject 
oEmail.HTMLBody = message
 
oEmail.Send
 
Set oConf = Nothing
Set oEmail = Nothing
