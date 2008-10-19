<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<%
if session("username")="" then
  Response.Redirect "login.asp"
end if

session("menu")="9"
Response.Redirect "main.asp"
%>