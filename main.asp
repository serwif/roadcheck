<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>
<%
dim tjbb
if not isempty(request("tjbb")) then
    tjbb = request("tjbb")
else
    tjbb = "fmc"
end if
%>
<HTML>
<HEAD>
<TITLE>三明市公路通行费管理系统</TITLE>
</HEAD>
<%if tjbb="dl" then
  Response.Redirect "login.asp"
elseif tjbb="zx" then
  Response.Redirect "logout.asp"
else
%>
<frameset rows="85,*" cols="*"  frameborder="0" border="0">
  <frame name=up src=top.asp?tjbb=<%=tjbb%> scrolling=no> 
  <%if tjbb="xtsz" then%>
    <frameset cols="132,*" rows="*">
      <frame name=left src=menu7.asp> 
      <frame name=right src=aboutthis.asp>  
    </frameset>
  <%else%>
    <frameset >
      <%if tjbb="fmc" then%>
        <frame name=right src=searchfmc.asp>    
      <%elseif tjbb="fgw" then%>
        <frame name=right src=searchqczlcs.asp>    
      <%elseif tjbb="jtb" then%>
        <frame name=right src=searchzfsccx.asp>    
      <%end if%>
    </frameset>
  <%end if%>
</frameset>
<%end if%>
<noframes>你的浏览器不支持FRAME！！！
</noframes> 
</HTML>