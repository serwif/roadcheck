<!--#include file=conn.asp -->
<!--#include file="const.asp"-->
<!--#include file="inc/char.inc"-->
<html>
<head>
<title>公告</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="style.asp"-->
<LINK href=style.css rel=stylesheet>
</head>

<body bgcolor="#FFFFFF" text="#000000">


<table border="0" cellspacing="0" cellpadding="0" width="100%" align="center" bgcolor="<%=RightBColor%>">
  <tr> 
    <td bgcolor="<%=RightCColor%>"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="8" align="center" style="TABLE-LAYOUT: fixed">
        <tr> 
          <td style="WORD-WRAP: break-word"> 
            <table border="0" width="100%" cellspacing="0">
              <tr> 
                <%
sql="SELECT * FROM Announce where popup="&true&" and Ischecked="&true&" order by id desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
%>
                <td width="100%" align="center">尚无任何公告</td>
                <%
else
%>
                <td width="100%"> 
                  <%
do while not rs.EOF
%>
                  <table border="0" width="100%" cellspacing="2" cellpadding="2">
                    <%	if not isnull(rs("Title")) then%>
                    <TR> 
                      <td width="100%" align=center height=57><font color=<%=AlertFColor%> size="+1"><b><%=rs("Title")%></b></font></td>
                    </TR>
                    <%
end if

%>
                    <TR> 
                      <td width="100%"><%=htmlencode2(rs("Content"))%></td>
                    </TR>
                  </table>
                  <%
rs.movenext
loop%>
                </td>
                <%
end if
rs.close
%>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%set rs=nothing
conn.close
set conn=nothing
%>

</body>
</html>
