<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function_title.asp" -->
<%
keyeord=request("keyword")
newsid=cint(request("newsid"))
javastr="<table border=0 cellspacing=0 cellpadding=0 width=""100%"" style=""BORDER-LEFT: "&CenterBColor&" 1px double; BORDER-RIGHT: "&CenterBColor&" 1px double; BORDER-BOTTOM: "&CenterBColor&" 1px double"">"
javastr=javastr+"<tr><TD bgcolor="""&CenterBColor&""" background="""&weburl&CenterBImg&""" HEIGHT=1></TD></tr>"
javastr=javastr+"<tr><td width=""100%"" bgcolor="""&CenterTColor&"""  background="""&weburl&CenterTImg&""" height=""18"" class=maintitle>&nbsp;相关信息："&about&"</td></tr>"
javastr=javastr+"<tr><TD bgcolor="""&CenterBColor&""" background="""&weburl&CenterBImg&""" HEIGHT=1></TD></tr>"
javastr=javastr+"<tr><td bgcolor="""&CenterCColor&""" background="""&weburl&CenterCImg&""">"
javastr=javastr+"<table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0""><tr height=114><td valign=top>"
if about<>"" then
	sql="select top 5 title,image,updatetime,newsid from news where checked="&true&" and about like '%" & keyeord & "%' and NewsID<>" & newsid &" order by newsid desc"
	set rs=conn.execute(sql)	
		do while not rs.eof	
			javastr=javastr+shownewf
			if ShowNewsModelRight=1 then 
				javastr=javastr+ showTitle("MainContentS",30)
			else
				javastr=javastr+ showTitle("MainContentS",44)
			end if
		javastr=javastr+ showImg
		javastr=javastr+ showTime
		javastr=javastr+ "<br>"
		rs.movenext
		loop
		javastr=javastr+ "<div align=""right""><a Class=""MainMore"" href="""&weburl&"showsearch.asp?keyword=" & keyeord &""" target=_self>>>更多</a>&nbsp;</div>"
	rs.close		
else
		javastr=javastr+ "&nbsp;尚无信息"
end if
javastr=javastr+ "</td></tr></table></td></tr></table>"
response.write ("document.write('"&javastr&"');")
response.end
set rs=nothing
conn.close
set conn=nothing		
%>