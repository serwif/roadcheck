<%PageName="Search"%>
<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<%
ifBigClassName=Request.form("BigClassName")
ifSmallClassName=Request.form("SmallClassName")
if ifSmallClassName<>"" then a4="��"

if ifBigClassName<>"" then
	a2="��"
	a5="��"
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
alert("�������ѯ�ؼ��֣�")
</script>
<%
Response.End
end if

findword=" and (title like '%"&keyword&"%' or content like '%"&keyword&"%' or author like '%"&keyword&"%' or original like '%"&keyword&"%') "
if request("action")="title" then
findword=" and title like '%"&keyword&"%' "
aa="����"
end if
if request("action")="content" then
findword=" and content like '%"&keyword&"%' "
aa="����"
end if
if request("action")="author" then
findword=" and author like '%"&keyword&"%' "
aa="����"
end if
if request("action")="original" then
findword=" and original like '%"&keyword&"%' "
aa="��Դ"
end if

if request("action")<>"" then a1="��"							
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
		if InStr(S_Key,"=")<>0 or InStr(S_Key,"`")<>0 or InStr(S_Key,"'")<>0 or InStr(S_Key," ")<>0 or InStr(S_Key,"��")<>0 or InStr(S_Key,"'")<>0 or InStr(S_Key,chr(34))<>0 or InStr(S_Key,"\")<>0 or InStr(S_Key,",")<>0 or InStr(S_Key,"<")<>0 or InStr(S_Key,">")<>0 then 
		Response.Redirect "error.htm" 
		End If 
		lngLenKey=Len(S_Key) 
		if lngLenKey>1 then
			For i=1 To lngLenKey-(lngSubKey-1) 
			strSubKey(i)=Mid(S_Key,i,lngSubKey) 
			if request("action")="" or request("action")="content" and request("Submit")="ģ" then strNew1=strNew1 & " or content like '%" & strSubKey(i) & "%'" 
			if request("action")="" or request("action")="title" and request("Submit")="ģ" then strNew2=strNew2 & " or title like '%" & strSubKey(i) & "%'" 
			Next 
		End if
	sql="Select * from news where checked="&true&" "& sql1 & findword & strNew1 & strNew2
	rs.Open sql,conn,1,1	
	If rs.BOF And rs.EOF Then 
		%> 
		      <p>&nbsp;</p>
              <p><font color="#FF0000"><b>δ�ҵ��κν��</b></font></p>
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
		                <td bgcolor=#3366cc nowrap><font color=#ffffff size="2">��<%=a1%><%=aa%><%=a2%><%=ifBigClassName%><%=a4%><%=ifSmallClassName%><%=a5%>�����йء�<b><%=S_Key%></b>������</font></td>
		                <td bgcolor=#3366cc align=right nowrap> <font size="-1" color="#ffffff">�����ǲ�ѯ��� 
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
	if request("Submit")<>"ģ" then
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
      <font color="#FF00FF">[ͼ]</font>
      <%end if%>
    </td>
  </tr>
  <tr>
    <td><%=strContent%><b>...</b></td>
  </tr>
  <tr>
                  <td><font color=green> 
                    <%if rs("Original")<>"" then%>
                    [��Դ��<%=replace(rs("Original"),S_Key,"<font color=red>"&S_Key&"</font>")%>] 
                    <%end if%>
                    <%if rs("Author")<>"" then%>
                    [���ߣ�<%=replace(rs("Author"),S_Key,"<font color=red>"&S_Key&"</font>")%>] 
                    <%end if%>
                    [����ʱ�䣺<%=rs("updatetime")%>]</font></td>
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
		if InStr(strKey,"=")<>0 or InStr(strKey,"`")<>0 or InStr(strKey,"'")<>0 or InStr(strKey," ")<>0 or InStr(strKey,"��")<>0 or InStr(strKey,"'")<>0 or InStr(strKey,chr(34))<>0 or InStr(strKey,"\")<>0 or InStr(strKey,",")<>0 or InStr(strKey,"<")<>0 or InStr(strKey,">")<>0 then 
		Response.Redirect "error.htm" 
		End If 
		lngLenKey=Len(strKey) 
		Select Case lngLenKey 
		Case 0 '��Ϊ�մ���ת������ҳ 
		Response.Redirect "error.htm" 
		Case 1 '������Ϊ1�������κ�ֵ 
		strNew1="" 
		strNew2="" 
		Case Else '�����ȴ���1������ַ������ַ���ʼ��ѭ��ȡ����Ϊ2�����ַ�����Ϊ��ѯ���� 
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
                �������� <font color="<%=AlertFColor%>"><b><%=totalnumber%></b></font> 
                ��
                <%if CurrentPage<2 then%>
                &nbsp;��ҳ &nbsp;��һҳ&nbsp; 
                <%else%>
                &nbsp<a href="<%=filename%>?page=1&<%=url%>">��ҳ</a>&nbsp; <a href="<%=filename%>?page=<%=CurrentPage-1%>&<%=url%>">��һҳ</a>&nbsp; 
                <%
		end if
		if n-currentpage<1 then
		%>
                ��һҳ&nbsp;&nbsp;ĩҳ 
                <%else%>
                <a href="<%=filename%>?page=<%=CurrentPage+1%>&<%=url%>">��һҳ</a> 
                &nbsp; <a href="<%=filename%>?page=<%=n%>&<%=url%>">ĩҳ</a> 
                <%end if%>
                &nbsp;ҳ�Σ�<strong><font color="<%=AlertFColor%>"><%=CurrentPage%>/<%=n%></font></strong>ҳ 
                ת���� 
                <select name="page" size="1" onchange="javascript:submit()">
		<%for i = 1 to n%>
		<option value="<%=i%>" <%if cint(CurrentPage)=cint(i) then%> selected <%end if%>>��<%=i%>ҳ</option>
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
