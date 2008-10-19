<!--#include file="conn.asp" -->
<!--#include file="const.asp" -->
<!--#include file="function_title.asp" -->
<%
newsid=int(request("newsid"))
if request("username")<>"" then username_news_count(request("username"))
if request("move")<>"" then movenews(newsid)

conn.close
set conn=nothing

function username_news_count(strusername)
sql="select newsid from News where username='"&strusername&"'"
rs.open sql,conn,1,1
response.write ("document.write('"&rs.RecordCount&"');")
response.end
rs.Close
set rs=nothing
end function

function movenews(strnewsid)
if request("move")="next" then 
strnewsid=strnewsid+1
else
strnewsid=strnewsid-1
end if
	sql="select newsid,title,image,updatetime from News where checked="&true&" and NewsID=" & strnewsid
	rs.open sql,conn,1,1
	if not rs.EOF then
		javastr=showTitle("MainContentS",70)
		javastr=javastr+showImg
		javastr=javastr+showTime		
	else
		javastr="已经没有了"
	end if
	response.write ("document.write('"&javastr&"');")
	response.end
	rs.close
	set rs=nothing	
end function
%>