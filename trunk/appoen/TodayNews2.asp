<!--导读开始-->
<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function.asp" -->
<%
javastr=""
javastr=javastr+"<table width=98% border=0 align=center cellspacing=0 cellpadding=0 bgcolor="&CenterBColor&">"
if showboard=1 then javastr=javastr+InTable("middle2")
javastr=javastr+"<tr><td colspan=2 bgcolor="&CenterTColor&" height=18 background="""&CenterTImg&"""><table width=""100%""><tr><td valign=bottom class=MainTitle width=""100%"">"
javastr=javastr+"&nbsp;<img src=""images/cat.gif"">&nbsp;<b>最新导读</b>　"&now()
javastr=javastr+"</td></tr></table></td></tr>"
javastr=javastr+InTable("middle2")
javastr=javastr+"<tr><td width=130 bgcolor="&CenterCColor&" valign=top>"
javastr=javastr+"<table width=100% border=0 align=left cellspacing=1 cellpadding=0>"
javastr=javastr+"<tr>"
javastr=javastr+"<td bgcolor="&CenterCColor&">"
javastr=javastr+"<table border=0 cellspacing=0 cellpadding=4>"
sql="select top 2 * from News where (image>0 and checked=1) order by updatetime DESC"
rs.open sql,conn,1,1
n=1
while not rs.EOF
javastr=javastr+"<tr><td align=center>"
javastr=javastr+ImageFile(rs("NewsID"),1,125,95)
javastr=javastr+"</td></tr>"
javastr=javastr+"<tr><td align=center>"
javastr=javastr+ShowTitle("MainContentS",20)
if n=1 then javastr=javastr+"<br><br>"
javastr=javastr+"</td></tr>"				
n=n+1
rs.MoveNext
wend
rs.close

javastr=javastr+"</table></table></td>"
javastr=javastr+"<td bgcolor="&CenterCColor&" valign=top align=left>"
javastr=javastr+"<table width=100% border=0 cellspacing=1 cellpadding=0 align=left>"
javastr=javastr+"<tr>"
javastr=javastr+"<td bgcolor="&CenterCColor&" align=left>" 

sql="select top 14 "& NoContent &" from News where checked=1 order by updatetime DESC"
rs.Open sql,conn,1,1
n=1
while not rs.EOF
javastr=javastr+"&nbsp;·"&ShowTitle("MainContentS",40)&"<br>"
if n=7 then javastr=javastr+"<br>"
n=n+1
rs.MoveNext
wend
rs.close

set rs=nothing
conn.close
set conn=nothing

javastr=javastr+"</table></td></tr></table><br>"
response.write ("document.write('"&javastr&"')")
response.end
%>
<!--导读结束-->