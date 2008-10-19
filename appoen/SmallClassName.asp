<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<%
BigClassName=Request("BName")
if Request("BType")="" or Request("BType")="1" then 
	BigClassType=1
else
	BigClassType=0
end if 
sql="select SmallClassType,SmallClassName from SmallClass Where BigClassName='" & BigClassName &"' order by SmallClassID"
rs.open sql,conn,1,1
dim SmallClassCount
SmallClassCount=rs.RecordCount
if SmallClassCount>0 then
	dim ArraySmallClassName(100),ArraySmallClassType(100)
	for i=1 to rs.RecordCount
		ArraySmallClassType(i)=rs(0)
		ArraySmallClassName(i)=rs(1)
		rs.MoveNext
	next
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing
if SmallClassCount>1 then
	javastr="<table border=0 cellspacing=0 cellpadding=0 width=100% align=center bgcolor="&LeftBColor&">"
	javastr=javastr+ TTitle("left","相关小类")
	javastr=javastr+ "<tr>"&_
	"<td width=100% bgcolor="&LeftCColor&" background="""&LeftCImg&""">"&_
	"<table width=100% border=0 cellspacing=0 cellpadding=0>"
	
	for i=1 to SmallClassCount
	javastr=javastr+ "<tr><td align=center height=20>"
	dim b1,b2
	if ArraySmallClassName(i)=request("SmallClassName") then
	b1="<img src=""images/icon111.gif"" board=0><b>"
	b2="</b><img src=""images/icon112.gif"" board=0>"
	else
	b1=""
	b2=""
	end if
	dim S_BigClassType,S_SmallClassType
	if Request("BType")="" or cint(Request("BType"))=1 then
		S_BigClassType=""
	else
		S_BigClassType="&BType=0"
	end if
	if ArraySmallClassType(i)=1 then
		S_SmallClassType=""
	else
		S_SmallClassType="&SType=0"
	end if
	javastr=javastr+ b1&"<a Class=LeftMenu href=""SmallClass.asp?BName="& BigClassName & S_BigClassType & "&SName="& ArraySmallClassName(i) & S_SmallClassType &""">" &ArraySmallClassName(i) &"</a>"&b2
	javastr=javastr+ "</td></tr>"
	next
	javastr=javastr+ "</table></td></tr>"
	javastr=javastr+ InTable("bottoml")
	javastr=javastr+ "</table>"
	response.write ("document.write('"&javastr&"')")
end if
%>