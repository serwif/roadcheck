<%PageName="Search"%>
<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function.asp" -->
<%
request_BigClassName=Request.form("BigClassName")
request_SmallClassID=Request.form("SmallClassID")
if request_SmallClassID<>"" then
sql="select SmallClassName from SmallClass where SmallClassID="&request_SmallClassID
rs.open sql,conn,1,1
if not rs.eof then
request_SmallClassName=rs("SmallClassName")
a4="��"
end if
rs.close
else
a4=""
end if

if request_BigClassName="" then
a2=""
a5=""
else
a2="��"
a5="��"
end if
keyword=trim(Request("keyword"))
if not isempty(request("soft")) then
soft=request("soft")
a3=""
else
a3="������"
soft="NEWSID DESC"
end if

if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if

if  request("action")<>"" and (keyword="�޹ؼ���" or keyword="") then
%>
<script language=javascript>
history.back()
alert("�������ѯ�ؼ��֣�")
</script>
<%
Response.End
end if

if request("action")="" and keyword="�޹ؼ���" then findword=""
if request("action")="" and keyword<>"�޹ؼ���" then findword=" title like '%"&keyword&"%' or content like '%"&keyword&"%' or author like '%"&keyword&"%' or original like '%"&keyword&"%' "
if request("action")="title" then
findword=" title like '%"&keyword&"%' "
aa="����"
end if
if request("action")="content" then
findword=" content like '%"&keyword&"%' "
aa="����"
end if
if request("action")="author" then
findword=" author like '%"&keyword&"%' "
aa="����"
end if
if request("action")="original" then
findword=" original like '%"&keyword&"%' "
aa="��Դ"
end if

if request("action")="" then
a1=""
else
a1="��"
end if
		
if keyword<>"�޹ؼ���" or keyword<>"" then
keyword=Keyword
else
keyword="ȫ��"
end if		
					
if findword="" then
ifand=""
else
ifand=" and "
end if
htmltop(0)
%>
<table border="0" width="760" cellspacing="0" cellpadding="0" bgcolor="<%=MainBgcolor%>">
<tr> <%=OutTable("left")%>
<td align="center" valign="top" bgcolor="<%=LeftBgColor%>">

<script language="javascript" src=search.asp?align=w></script>

<p><b><font size="2"><%=a1%><%=aa%><%=a2%><%=request_BigClassName%><%=a4%><%=request_SmallClassName%><%=a5%><%=a3%><%=keyword%></font></b></p>
<table border="0" cellpadding="1" cellspacing="0" width="98%">
<tr>
<td width=100% height="10" valign="middle" align="center"> </td>
</tr>
<tr>
<td align="center">
<%		
if request_BigClassName><"" and request_SmallClassName><"" then sql="select "& NoContent &" from News where BigClassName='" & request_BigClassName &"' and SmallClassName='" & request_SmallClassName &"' and checked="&true&" "& ifand & findword & " order by "&soft

if request_BigClassName><"" and request_SmallClassName="" then sql="select "& NoContent &" from News where BigClassName='" & request_BigClassName &"' and checked="&true&" "& ifand & findword & " order by "&soft

if request_BigClassName="" then sql="select "& NoContent &" from News where checked="&true&" "&ifand & findword & " order by "&soft

rs.Open sql,conn,1,1
if rs.eof and rs.bof then
response.write "<p align='center'><br><b>û �� �� �� �� Ҫ �� �� ��</b><br><br></p>"
else
MaxPerPage=20   '��������
PageUrl="showsearch.asp"
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
<table border="1" cellspacing="0" cellpadding="2" class="TableLine" bordercolorlight="<%=MainBColor%>" bgcolor="<%=MainCColor%>" width="100%">
<tr bgcolor="<%=CenterTColor%>">

					
<td height="10" align="center" width="9%"><a class=noline href="showsearch.asp?soft=NewsID&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a>����ID<a class=noline href="showsearch.asp?soft=NewsID DESC&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a></td>

					
<td height="10" align="center" width="39%"><a class=noline href="showsearch.asp?soft=Title&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a>���±���<a class=noline href="showsearch.asp?soft=Title DESC&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a></td>
<td height="10" align="center" width="9%"> <a class=noline href="showsearch.asp?soft=author&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a>����<a class=noline href="showsearch.asp?soft=author DESC&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a></td>
                     					
<td height="10" align="center" width="8%"><a class=noline href="showsearch.asp?soft=original&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a>��Դ<a class=noline href="showsearch.asp?soft=original DESC&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a></td>
                     					
<td height="10" align="center" width="18%"><a class=noline href="showsearch.asp?soft=UpdateTime&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a>����ʱ��<a class=noline href="showsearch.asp?soft=UpdateTime DESC&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a></td>
                     					
<td height="10" align="center" width="8%"><a class=noline href="showsearch.asp?soft=Image&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a>ͼƬ<a class=noline href="showsearch.asp?soft=Image DESC&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a></td>
                     					
<td height="10" align="center" width="9%"><a class=noline href="showsearch.asp?soft=Click&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a>���<a class=noline href="showsearch.asp?soft=Click DESC&keyword=<%=Request("keyword")%>&BigClassName=<%=request_BigClassName%>&SmallClassName=<%=request_SmallClassName%>" title="��������">��</a></td>
</tr>
<%
do while not rs.eof
i=i+1
if rs("author")="" then
author="δ֪"
else
author=rs("author")
end if
if rs("original")="" then
original="δ֪"
else
original=rs("original")
end if
title=rs("title")		
%>
<tr>
<td align="center" width="9%"><%=rs("NewsID")%></td>
<td width="39%" style="WORD-WRAP: break-word">&nbsp;
<%
if instr(title,keyword)>0 then title=replace(title,keyword,"<font color=red>"&keyword&"</font>")
Response.Write "<a href='"&newsurl&"' target=_blank>"&title&"</a>"

%></td>
<td align="center" width="9%">
<%
if instr(author,keyword)>0 then author=replace(author,keyword,"<font color=red>"&keyword&"</font>")
Response.Write author
%>
</td>
<td align="center" width="8%"><a style=cursr:hand title="<%=original%>">
<%
if instr(original,keyword)>0 then original=replace(original,keyword,"<font color=red>"&keyword&"</font>")
Response.Write original
%>
</a></td>
<td align="center" width="18%"><%=rs("UpdateTime")%></td>
<td align="center" width="8%"><%=rs("image")%></td>
<td align="center" width="9%"><%=rs("click")%></td>
</tr>
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
url="soft="&soft&"&BigClassName=" & request_BigClassName &"&SmallClassName=" & request_SmallClassName & "&keyword=" & keyword
%>
<form method=Post action="<%=filename%>?<%=url%>">
�� <font color="<%=AlertFColor%>"><b><%=totalnumber%></b></font>
������
<%if CurrentPage<2 then%>
&nbsp;��ҳ &nbsp;��һҳ&nbsp;
<%else%>
&nbsp<a href="<%=filename%>?page=1&<%=url%>">��ҳ</a>&nbsp; <a href="<%=filename%>?page=<%=CurrentPage-1%>&<%=url%>">��һҳ</a>&nbsp;
<%
end if
if n-currentpage<1 then
%>
��һҳ &nbsp;ĩҳ
<%else%>
<a href="<%=filename%>?page=<%=CurrentPage+1%>&<%=url%>">��һҳ</a> &nbsp;
<a href="<%=filename%>?page=<%=n%>&<%=url%>">ĩҳ</a>
<%end if%>
&nbsp;ҳ�Σ�<strong><font color="<%=AlertFColor%>"><%=CurrentPage%>/<%=n%></font></strong>ҳ
ת����
<select name="page" size="1" onchange="javascript:submit()">
<%for i = 1 to n%>
<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>
<%next%>
</select>
</form>
<%end function%>
</td>
</tr>
</table>

</td><%=OutTable("right")%>
</tr>
</table>	
<%
function showSTitle(strClass,strMaxLen)
dim strSubject,strTrueSubject,m_bOverFlow,strSpaceBar,strTip,strTarget
strSubject = HTMLDecode(rs("title"))	
strTrueSubject = GetTrueLength(strSubject, strMaxLen, strSpaceBar)
m_bOverFlow = checkOverFlow(strSubject, strMaxLen)
if m_bOverFlow = True then
strTip = strSubject
else
strTip = ""
end if
strTarget="_blank"
if pagename="shownews" then strTarget="_self"
if strClass="" then strClass="MainContentS"
if instr(strTrueSubject,keyword)>0 then strTrueSubject=replace(strTrueSubject,keyword,"<font color=red>"&keyword&"</font>")
showstitle= "<a href='"&newsurl&"' title='"&strTip&"' target='"&strTarget&"'>"&strTrueSubject&"</a>"
end function
set rs=nothing
htmlend(0)
%>