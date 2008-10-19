<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<%if ShowNewImg=1 then%>
<!--#include file="function.asp" -->
<%
dim BigClassName,SmallClassName,SpecialName,IfClass,IfTxt,IfTime
BigClassName=Request("BigClassName")
SmallClassName=Request("SmallClassName")
SpecialName=Request("SpecialName")

if BigClassName<>"" then IfClass=" BigClassName='"&BigClassName&"' and "
if SmallClassName<>"" then IfClass=" SmallClassName='"&SmallClassName&"' and "
if BigClassName<>"" and SmallClassName<>"" then IfClass=" BigClassName='"&BigClassName&"' and SmallClassName='"&SmallClassName&"' and "
if SpecialName<>"" then IfClass=" SpecialName='"&SpecialName&"' and "

sql="select top 5 newsid,title,model,image from News where "& IfClass &" image>0 and checked="&true&" order by updatetime DESC"
rs.open sql,conn,1,1
if not rs.EOF then

dim javastr
javastr=""
javastr=javastr+"<table border=""0"" cellspacing=""0"" cellpadding=""0"" width="&TableWidth&" bgcolor="""&RightBColor&""">"
javastr=javastr+"<tr>"
javastr=javastr+OutTable("left")
javastr=javastr+"<td bgcolor="""&RightCColor&""" align=""center"" valign=""top"" style=""BORDER-bottom: "&out3Color&" 1px double"">"
javastr=javastr+"<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
javastr=javastr+"<tr>"

while not rs.EOF

javastr=javastr+"<td align=""center"" valign=""top"">"
javastr=javastr+"<table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""TABLE-LAYOUT: fixed""><td height=10></td></tr>"
javastr=javastr+"<tr><td align=""center"" class=noline>"
javastr=javastr+ImageFile(rs(0),1,120,100)
javastr=javastr+"</td></tr>"
javastr=javastr+"<tr><td style=""WORD-WRAP: break-word"" align=""center""><a class=noline href="""&NewsUrl&""" target=""_blank"">"&left(rs(1),24)&"</a></td></tr>"
javastr=javastr+"<tr><td height=6></td></tr>"
javastr=javastr+"</table>"
javastr=javastr+"</td>"

rs.MoveNext
wend

javastr=javastr+"</tr></table>"
javastr=javastr+"</td>"
javastr=javastr+OutTable("right")
javastr=javastr+"</tr>"
javastr=javastr+"</table>"
response.write ("document.write('"&javastr&"')")
response.end

end if
rs.close
end if
set rs=nothing
conn.close
set conn=nothing
%>