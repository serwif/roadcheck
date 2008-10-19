<%PageName="Search"%>
<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<%
ifBigClassName=Request.form("BigClassName")
ifSmallClassName=Request.form("SmallClassName")
if ifSmallClassName<>"" then a4="－"

if ifBigClassName<>"" then
	a2="从"
	a5="中"
end if
	
keyword=trim(Request("key"))

if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if

if  keyword="" then
%>
<script language=javascript>
history.back()
alert("请输入查询关键字！")
</script>
<%
Response.End
end if

findword=" and (title like '%"&keyword&"%' or content like '%"&keyword&"%' or author like '%"&keyword&"%' or original like '%"&keyword&"%') "
if request("action")="title" then
findword=" and title like '%"&keyword&"%' "
aa="标题"
end if
if request("action")="content" then
findword=" and content like '%"&keyword&"%' "
aa="内容"
end if
if request("action")="author" then
findword=" and author like '%"&keyword&"%' "
aa="作者"
end if
if request("action")="original" then
findword=" and original like '%"&keyword&"%' "
aa="来源"
end if

if request("action")<>"" then a1="按"							
htmltop(0)
%>
<table border="0" width="<%=TableWidth%>" cellspacing="0" cellpadding="0" bgcolor="<%=MainBgcolor%>">
<tr> 
<%=OutTable("left")%>
      <td align="center" valign="top" bgcolor="<%=LeftBgColor%>"> 
        <script language="javascript" src=search.asp?align=w></script>
<table border="0" cellpadding="1" cellspacing="0" width="98%">
<tr>
<td width=100% height="10" valign="middle" align="center"> </td>
</tr>
<tr>
<td align="center">
<%
dim sql1
if ifBigClassName<>"" then sql1=" and BigClassName='" & ifBigClassName &"'"
if ifSmallClassName<>"" then sql1=sql1&" and SmallClassName='" & ifSmallClassName &"'"
 
dim currentPage
if not isempty(request("page")) and request("page")<>"" then
currentPage=cint(request("page"))
else
currentPage=1
end if

function replacehtml(fString)
	fString = Replace(fString,"&nbsp;"," ")
	fString = Replace(fString,"<P>", "")
	fString = Replace(fString,"</P>", "")	
	fString = Replace(fString,"<BR>", "")
	fString = replace(fString,"<" ,"&lt;")
	fString = replace(fString,">" ,"&gt;")
	fString = replace(fString,"[[" ,"&lt;")
	fString = replace(fString,"]]" ,"&gt;")
	replacehtml=fString
end function

Dim S_Key 
S_Key = Trim(Request("key"))
If S_Key <>"" then 
		CONST lngSubKey=2 
		Dim lngLenKey, strNew1, strNew2, j,strSubKey(20)
		if InStr(S_Key,"=")<>0 or InStr(S_Key,"`")<>0 or InStr(S_Key,"'")<>0 or InStr(S_Key," ")<>0 or InStr(S_Key,"　")<>0 or InStr(S_Key,"'")<>0 or InStr(S_Key,chr(34))<>0 or InStr(S_Key,"\")<>0 or InStr(S_Key,",")<>0 or InStr(S_Key,"<")<>0 or InStr(S_Key,">")<>0 then 
		Response.Redirect "error.htm" 
		End If 
		lngLenKey=Len(S_Key) 
		if lngLenKey>1 then
			For i=1 To lngLenKey-(lngSubKey-1) 
			strSubKey(i)=Mid(S_Key,i,lngSubKey) 
			if request("action")="" or request("action")="content" and request("Submit")="模" then strNew1=strNew1 & " or content like '%" & strSubKey(i) & "%'" 
			if request("action")="" or request("action")="title" and request("Submit")="模" then strNew2=strNew2 & " or title like '%" & strSubKey(i) & "%'" 
			Next 
		End if
	sql="Select * from news where checked="&true&" "& sql1 & findword & strNew1 & strNew2
	rs.Open sql,conn,1,1	
	If rs.BOF And rs.EOF Then 
		%> 
		      <p>&nbsp;</p>
              <p><font color="#FF0000"><b>未找到任何结果</b></font></p>
              <p> 
                <% 
	Else 
		dim MaxPerPage,PageUrl,totalPut
		MaxPerPage=10
		PageUrl="showsearch.asp"
		totalPut=rs.recordcount
		%>
              </p>
              <table width=100% border=0 cellpadding=1 cellspacing=0 bgcolor=#3366cc>
		<tr>
		<td bgcolor=#3366cc nowrap>
		<table width=100% border=0 cellpadding=1 cellspacing=0 bgcolor=#3366cc>
		<tr>
		                <td bgcolor=#3366cc nowrap><font color=#ffffff size="2">已<%=a1%><%=aa%><%=a2%><%=ifBigClassName%><%=a4%><%=ifSmallClassName%><%=a5%>搜索有关“<b><%=S_Key%></b>”的项</font></td>
		                <td bgcolor=#3366cc align=right nowrap> <font size="-1" color="#ffffff">以下是查询结果 
                          </font> </td>
		</tr></table></td></tr></table><br><br>		
		<%
		if currentpage<1 then currentpage=1
		if (currentpage-1)*MaxPerPage>totalput then
				if (totalPut mod MaxPerPage)=0 then
				currentpage= totalPut \ MaxPerPage
			else
				currentpage= totalPut \ MaxPerPage + 1
			end if
		end if
		if currentPage=1 then
				showpage totalput,MaxPerPage,PageUrl
			showContent
			showpage totalput,MaxPerPage,PageUrl
		else
			if (currentPage-1)*MaxPerPage<totalPut then
			rs.move  (currentPage-1)*MaxPerPage
			dim bookmark
			bookmark=rs.bookmark
				showpage totalput,MaxPerPage,PageUrl
			showContent
			showpage totalput,MaxPerPage,PageUrl
			else
			currentPage=1
				showpage totalput,MaxPerPage,PageUrl
			showContent
			showpage totalput,MaxPerPage,PageUrl
			end if
		end if
	end if
	rs.close	

	sub showContent
	i=0
	do while not rs.eof
	i=i+1
	strContent=replacehtml(Left(rs("content"),150))
	'strcontent=rs("content")
	strtitle=rs("title")
	if request("Submit")<>"模" then
		strtitle=replace(strtitle,S_Key,"<font color=red>"&S_Key&"</font>")
		strContent=replace(strContent,S_Key,"<font color=red>"&S_Key&"</font>")
	else
		for j=1 To lngLenKey	'-(lngSubKey-1) 
		strSubKey(j)=Mid(S_Key,j,lngSubKey) 
		strContent=replace(strContent,strSubKey(j),"<font color=red>"&strSubKey(j)&"</font>")
		if request("action")="title" then strtitle=replace(strtitle,strSubKey(j),"<font color=red>"&strSubKey(j)&"</font>")
		next
	end if
	%>
  <table width="100%" border="0" cellspacing="3" cellpadding="0">
    <tr>
    <td>
	<!--[<b><font color="#990000"><%=rs("BigClassName")%></font></b>]--&gt;[<b><font color="#990000"><%=rs("smallclassname")%></font></b>]--&gt;<a href="shownews.asp?newsid=<%=rs("newsID")%>" target="_blank"><%=strtitle%></a>-->
	[<b><font color="#990000"><%=rs("BigClassName")%></font></b>]--&gt;[<b><font color="#990000"><%=rs("smallclassname")%></font></b>]--&gt;<a href="shownews.asp?id=<%=rs("newsID")%>" target="_blank"><%=strtitle%></a>
      <%if rs("image")>0 then%>
      <font color="#FF00FF">[图]</font>
      <%end if%>
    </td>
  </tr>
  <tr>
    <td><%=strContent%><b>...</b></td>
  </tr>
  <tr>
                  <td><font color=green> 
                    <%if rs("Original")<>"" then%>
                    [来源：<%=replace(rs("Original"),S_Key,"<font color=red>"&S_Key&"</font>")%>] 
                    <%end if%>
                    <%if rs("Author")<>"" then%>
                    [作者：<%=replace(rs("Author"),S_Key,"<font color=red>"&S_Key&"</font>")%>] 
                    <%end if%>
                    [发表时间：<%=rs("updatetime")%>]</font></td>
  </tr>
  <tr>
      
    <td height="20">&nbsp;</td>
  </tr>
</table>
	<% 
	if i>=MaxPerPage then exit do
	rs.MoveNext 
	loop
	end sub

	Function AutoKey(strKey,strContent) 
		CONST lngSubKey=2 
		Dim lngLenKey, strNew1, strNew2, i, strSubKey 
		if InStr(strKey,"=")<>0 or InStr(strKey,"`")<>0 or InStr(strKey,"'")<>0 or InStr(strKey," ")<>0 or InStr(strKey,"　")<>0 or InStr(strKey,"'")<>0 or InStr(strKey,chr(34))<>0 or InStr(strKey,"\")<>0 or InStr(strKey,",")<>0 or InStr(strKey,"<")<>0 or InStr(strKey,">")<>0 then 
		Response.Redirect "error.htm" 
		End If 
		lngLenKey=Len(strKey) 
		Select Case lngLenKey 
		Case 0 '若为空串，转到出错页 
		Response.Redirect "error.htm" 
		Case 1 '若长度为1，则不设任何值 
		strNew1="" 
		strNew2="" 
		Case Else '若长度大于1，则从字符串首字符开始，循环取长度为2的子字符串作为查询条件 
		For i=1 To lngLenKey-(lngSubKey-1) 
		strSubKey=Mid(strKey,i,lngSubKey) 
		if strContent<>"" then AutoKeyColor=replace(strContent,strSubKey,"<font color=red>"&strSubKey&"</font>")
		strNew1=strNew1 & " or content like '%" & strSubKey & "%'" 
		strNew2=strNew2 & " or title like '%" & strSubKey & "%'" 
		Next 
		End Select
		if strContent="" then AutoKey="Select * from news where checked="&true&" and (content like '%" & strKey & "%' or title like '%" & strKey & "%'" & strNew1 & strNew2 &")"
	End Function 	
	
	function showpage(totalnumber,maxperpage,filename)
		dim n,url
		if totalnumber mod maxperpage=0 then
		n= totalnumber \ maxperpage
		else
		n= totalnumber \ maxperpage+1
		end if
		url="Key="&Trim(Request("key"))
		%>
		<form method=Post action="<%=filename%>?<%=url%>"><center>
                共搜索到 <font color="<%=AlertFColor%>"><b><%=totalnumber%></b></font> 
                条
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
	<%end function
	
	Set rs=Nothing 
end if 
%>
</td>
</tr>
</table>
<br>
</td>
<%=OutTable("right")%>
</tr>
</table>
<%htmlend(0)%>
