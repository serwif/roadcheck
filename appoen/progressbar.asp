<%@EnableSessionState=False%>

<%

On Error Resume Next

Set theProgress = Server.CreateObject("ABCUpload4.XProgress")  '�����ϴ��������

theProgress.ID = Request.QueryString("ID")

'������������xml��ʽ���

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

 

