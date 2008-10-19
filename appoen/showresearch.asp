<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function.asp" -->
<%
if request("site")="right" then
jsresearch("right")
else
jsresearch("left")
end if
set rs=nothing
conn.close
set conn=nothing

Function jsresearch(strsite)
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
	javastr="document.write('"
	javastr=javastr+"<table width=100% border=0 cellspacing=0 cellpadding=0>"
	javastr=javastr+TTitle(strsite,"网站调查33")
	javastr=javastr+"<tr><td align=center height=40 bgcolor="&CColor&" background="""&CImg&""">"
	javastr=javastr+"<table border=0 width=100% cellspacing=2 cellpadding=2><tr>"
	set rs=conn.execute("SELECT * FROM research where Ischecked="&true&"")
	if not rs.eof then
		javastr=javastr+"<td><table border=0 width=100% cellspacing=2 cellpadding=2>"
		javastr=javastr+"<tr><td class="&TitleClass&">&nbsp;<img src=""images/research.gif"" border=0>&nbsp;"&rs("Title")&"</td></tr>"
		javastr=javastr+"<form action=""researchresult.asp?Type="" target=""newwindow"" method=post name=research><tr><td valign=top class="&TitleClass&">"		
		dim ischecked
		for i=1 to 8
			javastr=javastr+"');"&vbcrlf
			javastr=javastr+"document.write('"		
		if rs("Select"&i)<>"" then
		if i=1 then ischecked="checked"
		javastr=javastr+"<input style=""background-color: "&CColor&";border: 0"" "&ischecked&" name=Options type=radio value="&i&">"&i&"."&rs("Select"&i)&"<br>"
		end if		
		next
		javastr=javastr+"</td></tr><tr><td height=30 align=center>"
		javastr=javastr+"');"&vbcrlf
		javastr=javastr+"document.write('"		
		javastr=javastr+"<input type=submit value=""提交"" id=submit1 name=submit1 class=submit>&nbsp;"
		javastr=javastr+"');"&vbcrlf		
		javastr=javastr+"document.write(""<input onClick=open_window('researchresult.asp?Type=view','research','width=420,height=250') type=button value=结果 id=button1 name=button1 class=submit>"");"&vbcrlf		
		javastr=javastr+"document.write('"		
		javastr=javastr+"</td></tr></form></table></td>"
	else
		javastr=javastr+"<td width=100% align=center>尚　无　调　查</td>"		
	end if
	rs.close
	javastr=javastr+"</tr></table></td></tr>"
	if strsite="right" then 
		javastr=javastr+InTable("bottomr")
	else
		javastr=javastr+InTable("bottoml")
	end if
	javastr=javastr+"</table>"
	javastr=javastr+"');"	
	response.write javastr
	response.end		
End Function
%>