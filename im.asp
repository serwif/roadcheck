<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>
<!--#include file="fcommon.asp"-->

<%
dim conn_msg, rs, sql, errmsg, founderror, i

sub opendb()
  set conn_msg=server.createobject("ADODB.CONNECTION")
  if err.number<>0 then
    err.clear
    Response.Redirect "error.asp?errid=1"
  else
    conn_msg.open sysconstr
    if err then
      err.clear
      Response.Redirect "error.asp?errid=1"
    end if
  end if
end sub

sub closedb()
  conn_msg.Close
  set conn_msg=nothing
end sub
%>
<HTML>
<HEAD>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<meta http-equiv='refresh' content='15;url=im.asp'>
<TITLE>IM</TITLE>
<link rel="stylesheet" type="text/css" href="main.css">
</HEAD>

<body bgcolor="#d6d3ce" text="#000000" leftmargin="0" topmargin="0">
<%noRightClick()%>
<table width="100%" height="60" cellspacing="0" cellpadding="0" border="0" align="center">
  <tr><td height="5"></td></tr>
  <tr><td bgcolor=<%=lightskincolor()%>>
 

    
<%
  opendb()
  set rs=server.createobject("adodb.recordset")
  rs.Open "select * from im order by sendtime desc", conn_msg,1,1
  if rs.recordcount<>0 then
%>
    <table width="100%" border="0" cellspacing="1" cellpadding="0" align="left">
<%
        rs.movefirst
        while not rs.EOF 
%>
          <tr>
              <td width=280 align=left valign=top><font color=blue><b><%=rs("sender")%></b></font>于<font color=red><%=GetFriendlyDateFormat(rs("sendtime"))%></font>发布：</td>
              <td width=500 align=left><%=htmlencode(rs("content"),0)%></td>
          </tr>
          <%rs.MoveNext
        wend
  else
%>
    </table>
<%
  end if
  rs.Close
  set rs=nothing
  closedb()
%>

  </td></tr>
</table>
</body>
</HTML>

