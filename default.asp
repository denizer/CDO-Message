<%
On error resume next

Set Message = CreateObject("CDO.Message")

Message.Subject = "Test Mail"
Message.From = "" ' from email address
Message.To = "" ' recipient email address
Message.CC = ""
Message.TextBody = "Test email message"

Dim Configuration
Set Configuration = CreateObject("CDO.Configuration")
Configuration.Load -1 ' CDO Source Defaults
Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "" ' SMTP server
Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "" ' email username
Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "" ' email password
Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

Configuration.Fields.Update

Set Message.Configuration = Configuration
Message.Send

If Err.Number = 0 Then 
	response.write "Email sent"
Else
	response.write "Email non sent"
End If
	
response.write Err.description

%>

