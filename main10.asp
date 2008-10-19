<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>
<%
if session("username")="" then
  Response.Redirect "login.asp"
end if
%>
<HTML>
<HEAD>
<TITLE>每日一览</TITLE>
</HEAD>
<frameset rows="85,*" cols="*"  frameborder="0" border="0">
	<frame name=up src=top.asp?visitor=1 scrolling=no> 
	<frameset >
                <frame name=right src=searchmryl.asp?mode=1&visitor=1>    
	</frameset>
</frameset>

<noframes>你的浏览器不支持FRAME！！！
</noframes> 
</HTML>