<!--#include file="conn.asp"-->
<%
'---------------------����û�������-------------------------------
function checkin(s)
s=trim(s)
s=replace(s," ","&amp;nbsp;")
s=replace(s,"'","&amp;#39;")
s=replace(s,"""","&amp;quot;")
s=replace(s,"&lt;","&amp;lt;")
s=replace(s,"&gt;","&amp;gt;")
checkin=s
end function

'---------------------���������м�����Ա-------------------------------
function checkAdmin1
sql="select * from admin where username='"&Session("username")&"'and password='"&Session("password")&"'"
rs.open sql,conn,1,1
if rs.EOF then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "admin_login.asp"
end if
rs.close
if Session("IsAdmin")<>true then response.redirect "admin_login.asp"
end function

'---------------------���߼�������Ա-------------------------------
function checkAdmin2
sql="select * from admin where username='"&Session("username")&"'and password='"&Session("password")&"'"
rs.open sql,conn,1,1
if rs.EOF then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "admin_login.asp"
end if
rs.close
if Session("IsAdmin")<>true or (session("KEY")<>4 and session("KEY")<>5) then response.redirect "admin_login.asp"
end function

'---------------------��鳬��������Ա-------------------------------
function checkAdmin3
sql="select * from admin where username='"&Session("username")&"'and password='"&Session("password")&"'"
rs.open sql,conn,1,1
if rs.EOF then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "admin_login.asp"
end if
rs.close
if Session("IsAdmin")<>true or session("KEY")<>5 then response.redirect "admin_login.asp"
end function

'---------------------����û�Email-------------------------------
function IsValidEmail(email)
IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
IsValidEmail = false
exit function
end if
for each name in names
if Len(name) <= 0 then
IsValidEmail = false
exit function
end if
for i = 1 to Len(name)
c = Lcase(Mid(name, i, 1))
if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
IsValidEmail = false
exit function
end if
next
if Left(name, 1) = "." or Right(name, 1) = "." then
IsValidEmail = false
exit function
end if
next
if InStr(names(1), ".") <= 0 then
IsValidEmail = false
exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
IsValidEmail = false
exit function
end if
if InStr(email, "..") > 0 then
IsValidEmail = false
end if
end function
'---------------------�������-------------------------------
sub error()
%>
<!--#include file="conn.asp"-->
<!--#include file="style.asp"-->
<>
<head>
<title>������ʾ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>BODY {
FONT-SIZE: 9pt
}
.body {
FONT-SIZE: 9pt
}
</STYLE>
<SCRIPT language=JavaScript>
<!--

function MM_swapImgRestore() { //v3.0
var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v3.0
var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_swapImage() { //v3.0
var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</SCRIPT>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0"onload="MM_preloadImages('../images/err_help2.gif','../images/err_close2.gif','../images/err_but2.gif')">
<TABLE height="100%" cellSpacing=0 cellPadding=0 width="100%" border=0>
<TBODY>
<TR>
<TD height="80%">
<TABLE class=body cellSpacing=0 cellPadding=0 width=400 align=center background=../images/err_bg1.gif border=0>
<TBODY>
<TR>
<TD width=10 height=23><IMG height=23 src="../images/err1.gif" width=24 border=0></TD>
<TD width=348 height=23>&nbsp;<FONT face="Arial, Helvetica, sans-serif" color=#000000>ERROR - ��������</FONT></TD>
<TD vAlign=baseline width=37 height=23 align=right><A onmouseover="MM_swapImage('close','','../images/err_close2.gif',1)" onmouseout=MM_swapImgRestore() href="javascript:window.close()"><IMG height=18 src="../images/err_close1.gif" width=15 border=0 name=close></A></TD>
<TD width=5 height=23><IMG height=23 src="../images/err2.gif" width=5 border=0 name=errr1_c4></TD>
</TR>
</TBODY>
</TABLE>
<TABLE class=body cellSpacing=0 cellPadding=0 width=400 align=center border=0>
<TBODY>
<TR vAlign=bottom>
<TD background=../images/err_bg.gif height=120>
<BLOCKQUOTE>
<DIV id=base>
<br>��������Ŀ���ԭ��
<ul>				
<%=errmsg%></ul>
</DIV>

</BLOCKQUOTE>
<P align=center><A onmouseover="MM_swapImage('back','','../images/err_but2.gif',1)" onmouseout=MM_swapImgRestore() href="javascript:history.go(-1)"><IMG height=20 src="../images/err_but1.gif" width=73 border=0 name=back></A>
</TD>
</TR>
<TR>
<TD height=2><IMG height=5 src="../images/err_bot.gif" width=400></TD>
</TR>
</TBODY>
</TABLE>
</TD>
</TR>
</TBODY>
</TABLE>
</body>
</html>
<%
end sub

sub JMail
Dim JMail,SendMail
Set JMail=Server.CreateObject("JMail.SMTPMail")
JMail.Logging=True
JMail.Charset="gb2312"
JMail.ContentType = "text/html"
JMail.ServerAddress=SMTPServer
JMail.Sender=FromUserEmail
JMail.Subject=topic
JMail.Body=mailbody
JMail.AddRecipient ForUserEmail
JMail.Priority=3
JMail.Execute 
Set JMail=nothing 
if err then 
	err.clear
	Response.Write "<center><b> ���Ź����Ѿ��򿪣������������֧�ַ��Ż��������ַ���󣬵����ż��޷�������</b>"
else
	Response.Write "<center><b> �ż��Ѿ�������</b>"
end if
end sub

sub CDONTS
Dim objCDO
Set objCDO = Server.CreateObject("CDONTS.NewMail")
'MailFormat �ʼ��ĸ�ʽ��0��Html 1�����ı���
'BodyFormat ���ӵĸ�ʽ��1�����������Զ���Ϊ�ɵ����0����mailformat=0ʱ���Ӳ��䣬�����Ϊ�ɵ����
'To         �ʼ����շ��ĵ��������ַ
'Importance �ʼ�����Ҫ�ԣ�0���� 1���� 2���ߣ�
'From       �ʼ����ͷ��ĵ��������ַ
'Subject    �ʼ�������
'Body       �ʼ�������
'Send       ��ɷ����ʼ��Ķ���
objCDO.To         = ForUserEmail 
objCDO.From       = FromUserEmail
objCDO.MailFormat = 0
objCDO.BodyFormat = 0
objCDO.Importance = 1
objCDO.Subject    = topic
objCDO.Body       = mailbody 
objCDO.Send
Set objCDO = Nothing
if err then 
	err.clear
	Response.Write "<center><b> ���Ź����Ѿ��򿪣������������֧�ַ��Ż��������ַ���󣬵����ż��޷�������</b>"
else
	Response.Write "<center><b> �ż��Ѿ�������</b>"
end if
end sub

sub ASPEmail
Set mailer=Server.CreateObject("ASPMAIL.ASPMailCtrl.1")  
recipient=ForUserEmail 
sender=FromUserEmail
subject=topic
message=mailbody
mailserver=SMTPServer
result=mailer.SendMail(mailserver, recipient, sender, subject, message)
if err then 
	err.clear
	Response.Write "<center><b> ���Ź����Ѿ��򿪣������������֧�ַ��Ż��������ַ���󣬵����ż��޷�������</b>"
else
	Response.Write "<center><b> �ż��Ѿ�������</b>"
end if
end sub
%>