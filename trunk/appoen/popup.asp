<!--#include file=conn.asp -->
<%
sql="select win_width,win_height from announce where popup="&true&" and ischecked="&true&""
rs.open sql,conn,1,1
if not rs.eof then
%>
<!-- 
window.open ('popuplist.asp', 'newwindow', 'height=<%=rs(0)%>, width=<%=rs(1)%>, top=0, left=0, toolbar=no, menubar=no, scrollbars=no, resizable=yes,location=no, status=no') 
--> 
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
