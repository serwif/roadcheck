<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="./main.css">
</head>

<body bgcolor="#d6d3ce" text="#000000" leftmargin="0" topmargin="0" onLoad="startclock()">

<%
session("skin")=Request.Cookies("skin")
if session("skin")="" then
  session("skin")="green"
end if
%>
<table width="100%" height="80" background="./images/skin/<%=session("skin")%>/head.jpg" cellspacing="0" cellpadding="0" border="0" align="center">
<tr><td height=10 align=right>
    <%if request("tjbb")="dl" then%>
      <%if session("username")="" then%>
        <span class="topfont1">��¼</span>
      <%else%>
        <a href="main.asp?tjbb=zx" target="_parent"><span class="topfont">ע��<%=session("name")%></span></a>       
      <%end if%>
      <a href="main.asp?tjbb=xtsz" target="_parent"><span class="topfont">ϵͳ����</span></a>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">������</span></a>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">���ҷ���ί����</span></a>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">��ͨ��������ί�����ֲ�</span></a>
    <%elseif request("tjbb")="zx" then%>
      <span class="topfont1">��¼</span>
      <a href="main.asp?tjbb=xtsz" target="_parent"><span class="topfont">ϵͳ����</span></a>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">������</span></a>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">���ҷ���ί����</span></a>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">��ͨ��������ί�����ֲ�</span></a>
    <%elseif request("tjbb")="xtsz" then%>
      <%if session("username")="" then%>
        <a href="main.asp?tjbb=dl" target="_parent"><span class="topfont">��¼</span></a>
      <%else%>
        <a href="main.asp?tjbb=zx" target="_parent"><span class="topfont">ע��<%=session("name")%></span></a>
      <%end if%>
      <span class="topfont1">ϵͳ����</span>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">������</span></a>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">���ҷ���ί����</span></a>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">��ͨ��������ί�����ֲ�</span></a>
    <%elseif request("tjbb")="fmc" then%>
      <a href="./appoen/default.asp" target="_parent"><span class="topfont">���¹���</span></a>
      <span class="topfont1">������</span>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">���ҷ���ί����</span></a>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">��ͨ��������ί�����ֲ�</span></a>
    <%elseif request("tjbb")="fgw" then%>
      <a href="./appoen/default.asp" target="_parent"><span class="topfont">���¹���</span></a>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">������</span></a>
      <span class="topfont1">���ҷ���ί</span>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">��ͨ��������ί�����ֲ�</span></a>
    <%elseif request("tjbb")="jtb" then%>
      <a href="./appoen/default.asp" target="_parent"><span class="topfont">���¹���</span></a>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">������</span></a>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">���ҷ���ί����</span></a>
      <span class="topfont1">��ͨ��������ί�����ֲ�</span>
    <%end if%>
    &nbsp;&nbsp;
</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td height=20 align=right>
<marquee width="400" scrolldelay=120 onmouseover='this.stop()' onmouseout='this.start()'>
  <%
    'Response.Write "<font color=white>��ӭ�����������н�ͨ�ֹ�·ͨ�зѹ���ϵͳϵͳ��</font>"
	'response.write session("strip")
  %>
</marquee>
<!--#include file="clock.asp"-->
</td></tr>
</table>
<table width="100%" height="5" cellspacing="0" cellpadding="0" border="0" align="center"><tr><td></td></tr></table>
</body>
</html>