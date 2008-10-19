<%PageName="SpecialList"%>
<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function_title.asp" -->
<%
ifclass=" SpecialName<>'无'"
if not isempty(request("page")) and request("page")<>"" then
currentPage=cint(request("page"))
else
currentPage=1
end if
htmltop(0)
%>

<table border="0" width="<%=TableWidth%>" cellspacing="0" cellpadding="0" bgcolor="<%=CenterBgcolor%>" height="360">
<tr>
<%=OutTable("left")%>
<td width="160" align="center" valign="top" bgcolor="<%=LeftBgColor%>">
<script language="javascript" src=search.asp></script>
<%if showTxtTop<>"0" then%>
<script language="javascript" src="hottxt.asp"></script>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=LeftBColor%>" height="100%">
<tr>
<td bgcolor="<%=LeftCColor%>" background="<%=LeftCImg%>"></td>
</tr>
</table>
</td>
<%=InTable("left")%>
<td align=center valign=top>
<p>&nbsp;</p>
<p><b><font size="5">专  题  列  表</font></b></p>	
<table border="0" cellpadding="1" cellspacing="0" width="95%" align="center"  style="TABLE-LAYOUT: fixed">
<tr>
<td width=100% height="10" valign="middle" align="center"> </td>
</tr>
<tr>
<td>
<%
sql="select * from Special  order by SpecialID DESC "
rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'><br><b>暂 时 没 有 新 闻</b></p>"
else
MaxPerPage=MaxList
PageUrl="SpecialList.asp"
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
<%trline()%>			
<tr>
<%
do while not rs.eof
set rs1=server.createobject("adodb.recordset")
sql="SELECT newsid FROM news where SpecialName='"&rs("SpecialName") &"'"
rs1.Open sql,conn,1,1
SpecialnewsCount=rs1.RecordCount
rs1.close
set rs1=nothing
i=i+1
strMaxLen=40  '限制标题长度
strSubject = HTMLDecode(rs("SpecialName"))
strTrueSubject = GetTrueLength(strSubject, strMaxLen, strSpaceBar)
m_bOverFlow = checkOverFlow(strSubject, strMaxLen)
if m_bOverFlow = True then
strTip = strSubject
else
strTip = ""
end if	
%>
<tr>
<td height="20">
<ul><ul><ul><li><a Class="MainContentB" href="Special.asp?SpecialName=<%=rs("SpecialName")%>"><%=strSubject%></a>&nbsp;(共 <%=SpecialnewsCount%> 条)</li></ul></ul>
</ul></td></tr>
<%
if i>=MaxPerPage then exit do
rs.movenext
loop
%>
</table>
<%
end sub

function showpage(totalnumber,maxperpage,filename)
if totalnumber mod maxperpage=0 then
n= totalnumber \ maxperpage
else
n= totalnumber \ maxperpage+1
end if
%>
<form method=Post action="<%=filename%>"><center>
共 <font color="<%=AlertFColor%>"><b><%=totalnumber%></b></font>
个专题
<%if CurrentPage<2 then%>
&nbsp;首页 &nbsp;上一页&nbsp;
<%else%>
&nbsp<a href="<%=filename%>?page=1">首页</a>&nbsp; <a href="<%=filename%>?page=<%=CurrentPage-1%>">上一页</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
下一页&nbsp;&nbsp;末页
<%else%>
<a href="<%=filename%>?page=<%=CurrentPage+1%>">下一页</a> &nbsp;
<a href="<%=filename%>?page=<%=n%>">末页</a>
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
</table>
</td>
<%=OutTable("right")%>
</tr>
</table>	
<%set rs=nothing
htmlend(0)
%>