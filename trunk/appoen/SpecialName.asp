<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<%
if request("list")<>"" then MaxSpecialList=cint(request("list"))
if request("site")="right" then
jsspecialname("right")
else
jsspecialname("left")
end if
set rs=nothing
conn.close
set conn=nothing

Function jsspecialname(strsite)
	dim CImg,CColor,TitleClass,javastr
	if strsite="right" then
		CImg=RightCImg
		CColor=RightCColor
		TitleClass="RightContent"
	else
		CImg=LeftCImg	
		CColor=LeftCColor
		TitleClass="LeftContent"
	end if
	sql="select Top " & MaxSpecialList & " * from Special order by SpecialID desc"
	rs.Open sql,conn,1,1
	if not rs.EOF then	
		javastr=javastr+"<table border=0 cellspacing=0 cellpadding=0 width=100% align=center bgcolor="&CColor&" background="""&CImg&""">"
		javastr=javastr+TTitle("left","专题文章")
		javastr=javastr+"<tr><td><table width=100% border=0 cellspacing=0 cellpadding=0>"
		while not rs.EOF
			dim b3,b4
			if rs(2)=request("SpecialName") then
				b3="<img src=""images/icon111.gif"" board=0><b>"
				b4="</b><img src=""images/icon112.gif"" board=0>"
			end if
			javastr=javastr+"<tr><td width=100% height=22 align=center>"&b3&"<a Class=LeftMenu href=""Special.asp?Name="&rs(2)&""">"&rs(2)&"</a>"&b4&"</td></tr>"
			rs.MoveNext
		wend
		javastr=javastr+"<tr><td valign=middle height=21 align=right><a Class=LeftMore href=""speciallist.asp""><img src=""images/more9.gif"" border=0 alt=""更多专题""></a>&nbsp;</td></tr></table></td></tr>"
		if strsite="right" then 
			javastr=javastr+InTable("bottomr")
		else
			javastr=javastr+InTable("bottoml")
		end if
		javastr=javastr+"</table>"
	end if
	rs.Close
	response.write ("document.write('"&javastr&"');")
	response.end		
End Function
%>