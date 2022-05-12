NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
set Email = CreateObject("CDO.Message")
Email.From = "send-xxx@example.com"
Email.To = "receive-xxx@example.com"
Email.Subject = "subject test"
text_file="C:\text.txt"
attachment="C:\file.txt"

Set fso=CreateObject("Scripting.FileSystemObject")
Set myfile=fso.OpenTextFile(text_file,1,Ture)
text_body=myfile.readall
myfile.Close
Email.Textbody = text_body
Email.AddAttachment attachment
with Email.Configuration.Fields
.Item(NameSpace&"sendusing") = 2
.Item(NameSpace&"smtpserver") = "smtp.163.com" 
.Item(NameSpace&"smtpserverport") = 25
.Item(NameSpace&"smtpauthenticate") = 1
.Item(NameSpace&"sendusername") = "send-xxx"  
.Item(NameSpace&"sendpassword") = "password" 
.Update
end with
Email.Send
Set Email=Nothing
msgbox "E-mail send OK!" ,, "info"