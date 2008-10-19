<%@EnableSessionState=False%>

<%

On Error Resume Next

Set theProgress = Server.CreateObject("ABCUpload4.XProgress")  '创建上传组件对象

theProgress.ID = Request.QueryString("ID")

'将返回数据以xml格式输出

%>

<?xml version="1.0" encoding="gb2312" ?>

<plan>

       <PercentDone><%=theProgress.PercentDone%></PercentDone>

       <min><%=Int(theProgress.SecondsLeft/60)%></min>

       <secs><%=theProgress.SecondsLeft Mod 60%></secs>

       <BytesDone><%=Round(theProgress.BytesDone / 1024, 1)%></BytesDone>

       <BytesTotal><%=Round(theProgress.BytesTotal / 1024, 1)%></BytesTotal>

       <BytesPerSecond><%=Round(theProgress.BytesPerSecond/1024, 1)%></BytesPerSecond>

       <Information><%=theProgress.Note%></Information>

</plan>

 

