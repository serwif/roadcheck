<!--#include file = conn.asp-->
<%
BigClassName=request.querystring("BigClassName")
SmallClassName=request.querystring("SmallClassName")
fname=request.querystring("fname")
newsid=cint(request.querystring("newsid"))
if fname<>"" then
	sql = "select click from News where fname='"&fname&"'"
	rs.open sql,conn,1,1
	%>
	   document.write(<%=rs(0)%>)
	<%
	  rs.close
	  set rs=nothing
end if
if newsid<>0 then
	sql = "select click from News where newsid="&newsid
	rs.open sql,conn,1,1
	%>
   javastr="<%=rs(0)%>"
   document.write(javastr)	
	<%
	  rs.close
	  set rs=nothing
end if
if BigClassName<>"" then
	sql = "select newsid from News where BigClassName='"&BigClassName&"'"
	rs.open sql,conn,1,1
	%>
	  document.write('<%=rs.RecordCount%>)
	<%
	  rs.close
	  set rs=nothing
end if
if SmallClassName<>"" then
	sql = "select newsid from News where SmallClassName='"&SmallClassName&"'"
	rs.open sql,conn,1,1
	%>
	  document.write(<%=rs.RecordCount%>)
	<%
	  rs.close
	  set rs=nothing
end if
%>
