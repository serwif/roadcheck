<%
Response.Write "<table border=0 cellspacing=0 cellpadding=0 width=100% align=center bgcolor="""&RightBColor&""">"
Response.Write TTitle("right","站点推荐")
Response.Write "<tr>"&_
"<td bgcolor="""&RightCColor&""" background="""&RightCImg&""" height=18>"&_
"<table width=95% border=0 cellspacing=0 cellpadding=0 align=center>"

sql="SELECT * FROM FriendSite where (LogoUrl is not Null) and IsGood="&true&" and IsOK="&true&""
rs.open sql,conn,1,1
if not Rs.eof then
do while not rs.eof

Response.Write "<tr><td width=100% align=center height=35>"&_
		  "<a href="""&rs("SiteUrl")&""" target=""_blank"">"&_
		  "<img src='"&rs("LogoUrl")&"' border=0 width=88 height=31 alt='名称："&rs("SiteName")&"&#13;站长："&rs("SiteAdmin")&"&#13;&#10;地址："&rs("SiteUrl")&"&#13;&#10;简介："&rs("SiteIntro")&"'>"&_
		  "</a></td></tr>"

rs.movenext
loop
end if
rs.close

sql="SELECT * FROM FriendSite where (LogoUrl is Null) and IsGood="&true&" and IsOK="&true&""
rs.open sql,conn,1,1
if not Rs.eof then
Response.Write "<tr><td width=100% align=center><hr size=1></td></tr>"
do while not rs.eof
Response.Write "<tr><td width=100% align=center><img src=""images/FriendSite.gif"" border=0><a href='"&rs("SiteUrl")&"' target=""_blank"" title='名称："&rs("SiteName")&"&#13;&#10;站长："&rs("SiteAdmin")&"&#13;&#10;地址："&rs("SiteUrl")&"&#13;&#10;简介："&rs("SiteIntro")&"'>"&rs("siteName")&"</a></td></tr>"
rs.movenext
loop
end if
rs.close
Response.Write "</table></td></tr>"
Response.Write "</table>"
%>