<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="inc/ubb.inc"-->
<!--#include file="admin/admin_function.asp"-->
<%
NewsID=request("NewsID")
sql="select * from News where checked="&true&" and NewsID=" & NewsID
rs.open sql,conn,1,1
if not rs.eof then
	fname=makefilename(rs("updatetime"))
	UpdateTime=rs("UpdateTime")
	username=rs("username")
	Original=rs("Original")
	Author=rs("Author")
	title=rs("title")
	titleurl=rs("titleurl")
	about=rs("about")
	image=rs("image")
	model=rs("model")
	bigclassname=rs("bigclassname")
	smallclassname=rs("smallclassname")
	specialname=rs("specialname")
	Content=rs("Content")
	rs.close
	set rs=nothing
	
	if titleurl="" or isnull(titleurl) then 
	WriteNews
	end if
	conn.close
	set conn=nothing
	
	response.redirect weburl&"html/"&mid(fname,1,4)&"/"&mid(fname,5,2)&"/"&fname&"-1.htm"
	else
	Response.Write("sadadasd")
end if
%>