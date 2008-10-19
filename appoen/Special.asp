<%PageName="Special"%>
<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function_title.asp" -->
<%
dim request_SpecialName
request_SpecialName=Request("Name")

dim currentPage
if not isempty(request("page")) and request("page")<>"" then
currentPage=cint(request("page"))
else
currentPage=1
end if
htmltop(0)
%>
<%if ShowNewImg=1 then%><script language="javascript" src="showimg.asp?SpecialName=<%=request_SpecialName%>"></script><%end if%>
<table border="0" width="<%=TableWidth%>" cellspacing="0" cellpadding="0" bgcolor="<%=MainBgcolor%>" height="360">
<tr>
<%=OutTable("left")%>
<td width="160" align="center" valign="top" bgcolor="<%=LeftBgColor%>">
<script language="javascript" src=SpecialName.asp?list=25></script>
<script language="javascript" src=search.asp></script>
<%if showTxtTop<>"0" then%>
<script language="javascript" src="hottxt.asp?SpecialName=<%=request_SpecialName%>"></script>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=LeftBColor%>" height="100%">
<tr>
<td bgcolor="<%=LeftCColor%>"></td>
</tr>
</table>
</td>
<%=InTable("left")%>
<td align=center valign=top>
<table border="0" cellpadding="1" cellspacing="0" width="95%" align="center"  style="TABLE-LAYOUT: fixed">
<tr>
<td width=100% valign="middle" align="center">
<p>&nbsp;</p><br>

<p><b><font size="5"><%=request_SpecialName%></font></b></p>
<br>
</td>
</tr>
<tr>
<td>
<!------------------------------------------------------------------------------------------------------------->
<%
sql="select newsid,title,model,updatetime,click,hot,goodnews,image from News where SpecialName='" & request_SpecialName & "' and checked="&true&" order by NewsID DESC "
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "<p align='center'><br><b>尚　无　内　容</b></p><br><br>"
else
	dim MaxPerPage,PageUrl,totalPut
	MaxPerPage=MaxList
	PageUrl="Special.asp"
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
			
sub showContent
i=0
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" class="TableLine" bordercolorlight="<%=MainBColor%>" bgcolor="<%=MainCColor%>"   style="TABLE-LAYOUT: fixed">
<%=trline()%>				
<tr>
<td width="100%" style="WORD-WRAP: break-word">
<%
do while not rs.eof
i=i+1
dim good
if rs("goodnews")=0 then
good=""
else
good="<font color='"&AlertFColor&"'>荐</font>"
end if
Response.Write shownewf
	if ShowSpecialModelRight=1 then 
		response.write ShowTitle("MainContentB",42)
	else
		response.write ShowTitle("MainContentB",60)
	end if
Response.Write showImg
Response.Write showTime
Response.Write showclick & "&nbsp;" & good
Response.Write "<br>"
if i>=MaxPerPage then exit do
rs.movenext
loop
%>
</td>
</tr>
</table>
<%
end sub

function showpage(totalnumber,maxperpage,filename)
dim n,url
if totalnumber mod maxperpage=0 then
n= totalnumber \ maxperpage
else
n= totalnumber \ maxperpage+1
end if
url="SpecialName=" & request_SpecialName									
%>
<form method=Post action="<%=filename%>?<%=url%>"><center>
共 <font color="<%=AlertFColor%>"><b><%=totalnumber%></b></font>
条文章
<%if CurrentPage<2 then%>
&nbsp;首页 &nbsp;上一页&nbsp;
<%else%>
&nbsp<a href="<%=filename%>?page=1&<%=url%>">首页</a>&nbsp; <a href="<%=filename%>?page=<%=CurrentPage-1%>&<%=url%>">上一页</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
下一页&nbsp;&nbsp;末页
<%else%>
<a href="<%=filename%>?page=<%=CurrentPage+1%>&<%=url%>">下一页</a>
&nbsp; <a href="<%=filename%>?page=<%=n%>&<%=url%>">末页</a>
<%end if%>
&nbsp;页次：<strong><font color="<%=AlertFColor%>"><%=CurrentPage%>/<%=n%></font></strong>页
转到：
<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>
<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>第<%=i%>页</option>
<%next%>
</select>
</form>
<%end function%>
</td>
</tr>
</table
><br>
</td>
<%if ShowSpecialModelRight=1 then%>
<%=InTable("right")%>
<!--显示右栏-->
<td bgcolor="<%=RightBgColor%>" width="160" align="center" valign="top">
<script language="javascript" src="goodnews.asp?SpecialName=<%=Request_SpecialName%>"></script>
<script language="javascript" src="hotimg.asp?SpecialName=<%=Request_SpecialName%>"></script>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
<tr>
<td bgcolor="<%=RightCColor%>" background="<%=RightCImg%>"></td>
</tr>
</table>
</td>
<%end if%>
<%=OutTable("right")%>
</tr>
</table>
<%set rs=nothing
htmlend(0)
%>