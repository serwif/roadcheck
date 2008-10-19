<%
'response.write 8
'response.flush
'option explicit
dim PageName
PageName="SmallClass"%>
<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function_title.asp" -->

<%
dim currentPage
if not isempty(request("page")) and request("page")<>"" then
	currentPage=cint(request("page"))
else
	currentPage=1
end if

dim request_BigClassName,request_BigClassType,request_BigTemplate,request_SmallClassName,request_SmallClassType
request_BigClassName=Request("BName")
if Request("BType")<>"" then
	request_BigClassType=0
else
	request_BigClassType=1
end if
request_BigTemplate=cint(Request("Template"))
request_SmallClassName=Request("SName")
if Request("SType")<>"" then
	request_SmallClassType=0
else
	request_SmallClassType=1
end if

dim sql
sql="select * from SmallClass Where BigClassName='" & request_BigClassName &"' order by SmallClassID"
rs.open sql,conn,1,1
dim SmallClassCount,i
SmallClassCount=rs.RecordCount
if SmallClassCount>0 then
	dim ArraySmallClassName(50),ArraySmallClassType(50)
	for i=1 to SmallClassCount
		ArraySmallClassType(i)=rs(2)
		ArraySmallClassName(i)=rs(4)
		rs.MoveNext
	next
end if
rs.Close

htmltop(0)
%>
<%if ShowNewImg=1 then%><script language="javascript" src="showimg.asp?BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>"></script><%end if%>
<table border="0" width="<%=TableWidth%>" cellspacing="0" cellpadding="0" bgcolor="<%=CenterBgcolor%>" height="360">
<tr>
<%=OutTable("left")%>
<td width="160" align="center" valign="top" bgcolor="<%=LeftBgColor%>">
<script language="javascript" src="SmallClassName.asp?BName=<%=request_BigClassName%>"></script>
<script language="javascript" src="specialname.asp?site=left"></script>
<script language="javascript" src=search.asp></script>

<%if showTxtTop<>"0" then%><script language="javascript" src="hottxt.asp?BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>"></script>
<%end if%>

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<tr>
<td bgcolor="<%=LeftCColor%>" background="<%=LeftCImg%>"></td>
</tr>
</table>
</td>
<%=InTable("left")%>
<td align=center valign=top bgcolor="<%=CenterBgColor%>">
<%
if request_SmallClassType=0 then%>
	<script language="javascript" src="ClassNews.asp?BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>"></script>
	<%
else
	%>
	<p>&nbsp;</p>
	<p><b><span class=READNEWSTITLE><%=request_SmallClassName%></span></b></p>
	<table border="0" cellpadding="0" cellspacing="0" width="95%" align="center"  style="TABLE-LAYOUT: fixed">
	<tr><td width=100% height="10"></td></tr>
	<%=trline()%>
	<tr><td>
	<%			
	sql="select "&NoContent&" from News where BigClassName='"&request_BigClassName&"' and SmallClassName='" & request_SmallClassName &"' and checked="&true&" order by updatetime DESC"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<p align='center'><br><b>尚　无　新　闻</b></p><br>"
	response.write request_BigClassName
	response.write "<br>"&request_SmallClassName
	else
		dim MaxPerPage,PageUrl,totalPut
		MaxPerPage=MaxList
		PageUrl="SmallClass.asp"
		totalPut=rs.recordcount
		if currentpage<1 then currentpage=1
		if (currentpage-1)*MaxPerPage>totalput then
			if (totalPut mod MaxPerPage)=0 then
				currentpage= totalPut \ MaxPerPage
			else
				currentpage= totalPut \ MaxPerPage + 1
			end if
		end if
		if currentPage=1 then
			'		showpage totalput,MaxPerPage,PageUrl
			showContent
			showpage totalput,MaxPerPage,PageUrl
		else
			if (currentPage-1)*MaxPerPage<totalPut then
				rs.move  (currentPage-1)*MaxPerPage
				dim bookmark
				bookmark=rs.bookmark
				'			showpage totalput,MaxPerPage,PageUrl
				showContent
				showpage totalput,MaxPerPage,PageUrl
			else
				currentPage=1
				'			showpage totalput,MaxPerPage,PageUrl
				showContent
				showpage totalput,MaxPerPage,PageUrl
			end if
		end if
	end if
	rs.close
	%>	
	</td>
	</tr>
	</table>
	<br>
	</td>
	<%if ShowSmallModelRight=1 then%>
		<%=InTable("right")%>
		<!--显示右栏-->
		<td bgcolor="<%=RightBgColor%>" width="160" align="center" valign="top">
		<script language="javascript" src="goodnews.asp?BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>"></script>
		<script language="javascript" src="hotimg.asp?BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>"></script>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
		<tr>
		<td bgcolor="<%=RightCColor%>" background="<%=RightCImg%>"></td>
		</tr>
		</table>
		</td>
	<%end if
end if	
response.write  OutTable("right")%>
</tr>
</table>

<%
sub showContent
	i=0
	response.write "<table border=0 width=""98%"" cellspacing=0 cellpadding=0 bgcolor="&MainCColor&" style=""TABLE-LAYOUT: fixed"" align=""center"">"&_
	"<tr>"&_
	"<td width=""100%"" style=""WORD-WRAP: break-word"">"
	do while not rs.eof
	i=i+1
	response.write Shownewf
	if ShowSmallModelRight=1 then 
		response.write ShowTitle("MainContentB",42)
	else
		response.write ShowTitle("MainContentB",60)
	end if	
	response.write ShowTime
	response.write ShowImg
	'response.write Shownew
	response.write Showclick
	response.write "<br>"
	if rs("goodnews")=true then Response.Write "&nbsp;<font color='"&AlertFColor&"'>荐</font>"				
	if i>=MaxPerPage then exit do
	rs.movenext
	loop
	response.write "</td></tr></table>"
end sub

function showpage(totalnumber,maxperpage,filename)
	dim n
	if totalnumber mod maxperpage=0 then
		n= totalnumber \ maxperpage
	else
		n= totalnumber \ maxperpage+1
	end if
	dim url
	url="BName=" & request_BigClassName &"&BType=0&SName=" & request_SmallClassName
	response.write "<form method=Post action="""&filename&"?"&url&"""><center>"&_
	"共 <font color="&AlertFColor&"><b>"&totalnumber&"</b></font> 条"
	if CurrentPage<2 then
		response.write "&nbsp;首页 &nbsp;上一页&nbsp;"
	else
		response.write "&nbsp<a href="""&filename&"?page=1&"&url&""">首页</a>&nbsp; <a href="""&filename&"?page="&CurrentPage-1&"&"&url&""">上一页</a>&nbsp;"
	end if
	if n-currentpage<1 then
		response.write "下一页&nbsp;&nbsp;末页"
	else
		response.write "<a href="""&filename&"?page="&CurrentPage+1&"&"&url&""">下一页</a>"
		response.write "&nbsp; <a href="""&filename&"?page="&n&"&"&url&""">末页</a>"
	end if
	response.write "&nbsp;页次：<strong><font color="&AlertFColor&">"&CurrentPage&"/"&n&"</font></strong>页"
	response.write "转到："
	response.write "<select name=""page"" size=""1"" onchange=""javascript:submit()"">"
		for i = 1 to n%>
			<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>第<%=i%>页</option>
		<%next
	response.write "</select></form>"
end function
	
set rs=nothing
htmlend(0)
%>