<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function_title.asp" -->
<%
newsid=cint(request("NewsID"))
specialname=request("specialname")
	javastr="<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" style=""BORDER-LEFT: "&CenterBColor&" 1px double; BORDER-RIGHT: "&CenterBColor&" 1px double; BORDER-BOTTOM: "&CenterBColor&" 1px double""><tr><TD bgcolor="""&CenterBColor&""" background="""&weburl&CenterBImg&""" HEIGHT=1></TD></tr><tr><td width=""100%"" bgcolor="""&CenterTColor&""" background="""&weburl&CenterTImg&""" height=""18"" Class=maintitle>&nbsp;所属专题："
	if specialname<>"无" and specialname<>"" then javastr=javastr+"<a Class=maintitle href="""&weburl&"Special.asp?Name="& specialname &""" target=_self>" & request("specialname") &"</a>"
	javastr=javastr+"</td></tr><tr><TD bgcolor="""&CenterBColor&""" background="""&weburl&CenterBImg&""" HEIGHT=1></TD></tr><tr><td bgcolor="""&CenterCColor&""" background="""&weburl&CenterCImg&"""><table width=""100%"" border=""0"" cellspacing=""5"" cellpadding=""0""><tr height=114><td width=""100%"" valign=top>"
if SpecialName="无" or SpecialName="" then
	javastr=javastr+"&nbsp;尚无信息"		
else	
	sql="select top 5 title,image,updatetime,newsid from News where (checked="&true&" and NewsID<>" & newsid & " and SpecialName='" & specialname & "') order by NewsID DESC"
	set rs=conn.execute(sql)
	if rs.EOF then
		javastr=javastr+"&nbsp;尚无信息"
	else
		while not rs.EOF
			javastr=javastr+shownewf
			if ShowNewsModelRight=1 then 
				javastr=javastr+showTitle("MainContentS",30)
			else
				javastr=javastr+showTitle("MainContentS",44)
			end if
			javastr=javastr+showImg
			javastr=javastr+showTime
			javastr=javastr+"<br>"
			rs.MoveNext
		wend
		javastr=javastr+"<div align=""right""><a Class=""MainMore"" href="""&weburl&"Special.asp?SpecialName="& specialname &""" target=_self>>>更多</a>&nbsp;</div>"
	rs.Close	
	end if	
end if	
	javastr=javastr+"</td></tr></table></td></tr></table>"	
	set rs=nothing
	conn.close
	set conn=nothing
	response.write ("document.write('"&javastr&"');")
	response.end	
%>