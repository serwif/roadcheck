<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<%
session("username")=""
session("password")=""
if session("visitor")=1 then
session("visitor")=0
Response.Redirect "loginmryl.asp"
else
Response.Redirect "login.asp"
end if
%>
