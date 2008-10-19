<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<%
if session("username")="" then
  Response.Redirect "login.asp"
end if

session("menu")="2"
Response.Redirect "main.asp"
%>