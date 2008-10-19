<!--#include file="conn.asp"-->
<%
ads= request.QueryString("ads")
url= request.QueryString("url")

if ads="" or url="" then
response.write "操作失败！请联系管理员"
else
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From ad where id="&ads, conn,1,3
rs("totalclick")=rs("totalclick")+1
todaycounter=rs("today")
if date()=rs("beginday") then
rs("today")=rs("today")+1
else
rs("yesterday")=todaycounter
rs("beginday")=date()
rs("today")=0
rs.update
end if
set rs=nothing
conn.close
set conn=nothing
%>
<meta http-equiv=refresh content="1; url=<%=url%>">
<%end if%>