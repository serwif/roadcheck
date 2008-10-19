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
        <span class="topfont1">登录</span>
      <%else%>
        <a href="main.asp?tjbb=zx" target="_parent"><span class="topfont">注销<%=session("name")%></span></a>       
      <%end if%>
      <a href="main.asp?tjbb=xtsz" target="_parent"><span class="topfont">系统设置</span></a>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">花名册</span></a>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">国家发改委公告</span></a>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">交通部、发改委征费手册</span></a>
    <%elseif request("tjbb")="zx" then%>
      <span class="topfont1">登录</span>
      <a href="main.asp?tjbb=xtsz" target="_parent"><span class="topfont">系统设置</span></a>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">花名册</span></a>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">国家发改委公告</span></a>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">交通部、发改委征费手册</span></a>
    <%elseif request("tjbb")="xtsz" then%>
      <%if session("username")="" then%>
        <a href="main.asp?tjbb=dl" target="_parent"><span class="topfont">登录</span></a>
      <%else%>
        <a href="main.asp?tjbb=zx" target="_parent"><span class="topfont">注销<%=session("name")%></span></a>
      <%end if%>
      <span class="topfont1">系统设置</span>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">花名册</span></a>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">国家发改委公告</span></a>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">交通部、发改委征费手册</span></a>
    <%elseif request("tjbb")="fmc" then%>
      <a href="./appoen/default.asp" target="_parent"><span class="topfont">文章管理</span></a>
      <span class="topfont1">花名册</span>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">国家发改委公告</span></a>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">交通部、发改委征费手册</span></a>
    <%elseif request("tjbb")="fgw" then%>
      <a href="./appoen/default.asp" target="_parent"><span class="topfont">文章管理</span></a>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">花名册</span></a>
      <span class="topfont1">国家发改委</span>
      <a href="main.asp?tjbb=jtb" target="_parent"><span class="topfont">交通部、发改委征费手册</span></a>
    <%elseif request("tjbb")="jtb" then%>
      <a href="./appoen/default.asp" target="_parent"><span class="topfont">文章管理</span></a>
      <a href="main.asp?tjbb=fmc" target="_parent"><span class="topfont">花名册</span></a>
      <a href="main.asp?tjbb=fgw" target="_parent"><span class="topfont">国家发改委公告</span></a>
      <span class="topfont1">交通部、发改委征费手册</span>
    <%end if%>
    &nbsp;&nbsp;
</td></tr>
<tr><td>&nbsp;</td></tr>
<tr><td height=20 align=right>
<marquee width="400" scrolldelay=120 onmouseover='this.stop()' onmouseout='this.start()'>
  <%
    'Response.Write "<font color=white>欢迎您访问三明市交通局公路通行费管理系统系统。</font>"
	'response.write session("strip")
  %>
</marquee>
<!--#include file="clock.asp"-->
</td></tr>
</table>
<table width="100%" height="5" cellspacing="0" cellpadding="0" border="0" align="center"><tr><td></td></tr></table>
</body>
</html>