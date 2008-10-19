<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function.asp" -->
<%
dim n,javastr
javastr=""
javastr=javastr+"<table width=98% border=0 align=center cellspacing=0 cellpadding=0 bgcolor="&CenterBColor&" height=100>"
if showboard=1 then	javastr=javastr+InTable("middle1")
    javastr=javastr+"<tr><td bgcolor="&CenterTColor&" height=18 background="""&CenterTImg&"""><table width=100% ><tr><td valign=bottom class=MainTitle width=100% >&nbsp;<img src=images/cat.gif>&nbsp;<b>最新导读</b>　"&now()&"</td></tr></table></td></tr>"
	javastr=javastr+InTable("middle1")
	javastr=javastr+"<tr>"
    javastr=javastr+"<td bgcolor="""&CenterCColor&""" background="""&CenterCImg&""" align=left valign=top>"
	javastr=javastr+"<table width=100% border=0 align=center cellspacing=1 cellpadding=0>"
    javastr=javastr+"<tr>"
javastr=javastr+"<td>"
javastr=javastr+"<table width=98% border=0 cellspacing=0 cellpadding=0 align=center>"
sql="select top 10 "& NoContent &" from News where checked="&true&" order by updatetime DESC"
rs.Open sql,conn,1,1
if not rs.eof then
	n=0
	set rs1=server.CreateObject("ADODB.RecordSet")
	while not rs.EOF
		n=n+1
		dim titlelen	
		rs1.Open "select top 1 BigClassType from BigClass where bigclassname='"&rs("BigClassName")&"'",conn,1,1
		if not rs1.eof then
		BigClassType=rs1(0)
		end if
		rs1.close
		if n mod 2=1  then javastr=javastr+ "<tr>"
		javastr=javastr+ "<td width=50% Class=MainContentS>"
		javastr=javastr+ "<b>[</b><a class=LeftMenu href=""BigClass.asp?BigClassName="&rs("BigClassName")&"&BigClassType="&BigClassType&""">"&rs("BigClassName")&"</a><b>]</b> "
		titlelen=30-wordlen(rs("BigClassName"))
		javastr=javastr+ ShowTitle("MainContentS",titlelen)
		javastr=javastr+ "</td>"
		if n mod 2=0  then javastr=javastr+ "</tr>"						
	rs.MoveNext
	wend
'else
	'javastr=javastr+ "<tr><td><center><b>尚　无　内　容</b></center></td></tr>"
end if
rs.close
set rs1=nothing
set rs=nothing
conn.close
set conn=nothing
javastr=javastr+ "</table></table></td></tr></table><BR>"
response.write ("document.write('"&javastr&"')")
response.end

function wordlen(strChinese)
dim lenTotal,strWord
lenTotal = 0
for i=1 to Len(strChinese)
strWord = mid(strChinese, i, 1)
if asc(strWord) < 0 or asc(strWord) > 127 then
lenTotal = lenTotal + 2
else
lenTotal = lenTotal + 1
end if
next
wordlen=lenTotal
end function
%>
<!--导读结束-->