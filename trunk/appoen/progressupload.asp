<%@EnableSessionState=False%> 

<%

Response.Expires = -10000

Server.ScriptTimeOut = 300

 

Set theForm = Server.CreateObject("ABCUpload4.XForm")

theForm.Overwrite = True

theForm.MaxUploadSize = 8000000

theForm.ID = Request.QueryString("ID")

Set theField = theForm("filefield1")(1)

If theField.FileExists Then

       theField.Save theField.FileName

End If

%>

 

<html>

<body>

���ͽ���

</body>

</html>
