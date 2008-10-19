<% Response.Buffer=True %>
<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<%
if request.QueryString("Type")="" then
if Request.ServerVariables("REMOTE_ADDR")=request.cookies("IPAddress") then
response.write"<SCRIPT language=JavaScript>alert('感谢您的支持，您已经投过票了，请勿重复投票，谢谢！');"
response.write"javascript:window.close();</SCRIPT>"
else
options=request.form("options")
response.cookies("IPAddress")=Request.ServerVariables("REMOTE_ADDR")
set rs=server.createobject("adodb.recordset")
sql="update research set answer"&options&"=answer"&options&"+1 where IsChecked="&true&""
rs.open sql,conn,1,3
set rs=nothing
end if
end if
%><head>
<title><%=webtitle%>调查结果</title>
<style type="text/css">
<!--
td {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size: 9pt;line-height:14pt; color: #623669;letter-spacing:1pt}
div {font-family:"Verdana", "Arial", "Helvetica", "sans-serif";font-size:9pt; color:#008080;letter-spacing:4;line-height:14pt}
span {font-family:"Verdana", "Arial", "Helvetica", "sans-serif";font-size:9pt; color:#008080;letter-spacing:2;line-height:14pt}
p {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size:9pt; color: #623669;line-height:14pt}
.alert {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size:9pt; line-height:14pt; color: red;letter-spacing:2}
a:link {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size:9pt; color:#AD70B6; text-decoration: none; }
a:visited {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size:9pt; color:#999999; text-decoration: none;}
a:hover {  font-family: "Verdana", "Arial", "Helvetica", "sans-serif"; font-size:9pt; color: #FF0000; text-decoration: underline;}
INPUT.text,INPUT.file,TEXTAREA{font-family:"Verdana", "Arial", "Helvetica", "sans-serif",宋体;color:#623669;background-color:#ffffff;border:1 solid #623669}
INPUT.Submit {height=20;border:1 solid #D6BDDE;font-family: "Verdana";BACKGROUND-COLOR:#F2EAF4; FONT-SIZE: 9pt ;font-color:#623669; PADDING-TOP: 1px}
BODY{FONT-FAMILY: 宋体; FONT-SIZE: 9pt;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;
SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;
SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
-->
</style></head>




<p align="center">
<div align="center"> 
  <table border="0" cellpadding="0" cellspacing="1" width="100%" height="48" bgcolor="#D6BDDE">
    <%
total=0
set rs=server.createobject("adodb.recordset")
sql="select * from research where IsChecked="&true&""
rs.open sql,conn,1,1
%>
    <tr> 
      <td width="100%" height="22" bgcolor="#F2EAF4"><font color="#000000">『</font><font color="#000073"><%=rs("Title")%></font><font color="#000000">』调查结果</font> 
        <font color="#000073">&nbsp; </font> </td>
    </tr>
    <tr> 
      <td width="100%" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
          <tr> 
            <td width="60%" height="22">&nbsp;调查项目</td>
            <td height="22">图形显示</td>
            <td align="center" height="22" width="30">票数</td>
          </tr>
          <%
for i=1 to 8
if rs("Select"&i)<>"" then
total=total+rs("answer"&i)
end if
next
for i=1 to 8
if rs("Select"&i)<>"" then
if total=0 then
answer=0
else
answer=(rs("answer"&i)/total)*100
end if
%>
          <tr> 
            <td width="60%">&nbsp;<%=i%>.<%=rs("select"&i)%></td>
            <td><img src=images/bar1.gif width=<%=int(answer)%> height=10><%=round(answer,3)%>%</td>
            <td align="center" width="30"><%=rs("answer"&i)%></td>
          </tr>
          <%
end if
next
%>
        </table>
      </td>
    </tr>
    <tr>
      <td width="100%" valign="top">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#F2EAF4">
          <tr>
            <td height="22" valign="bottom">&nbsp;开始时间：<%=rs("DateAndTime")%></td>
            <td align="right" height="22" valign="bottom">总票数：</td>
            <td width="30" height="22" align="center" valign="top"><%=total%></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</div>
<p align="center">【<a href="javascript:window.close()">关闭窗口</a>】
<% rs.close
set rs=nothing
conn.close
set conn=nothing %>