<%
option explicit
dim PageName
PageName="shownews"%>
<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function.asp" -->
<!--#include file="readimg.asp" -->
<%
dim newsID,Title,titleurl,about,Author,Original,AuthorR,OriginalR,UpdateTime,Content,hot,SpecialName,SmallClassName,click,image,username

NewsID=request("NewsID")

sql="select * from News where checked=1 and NewsID=" & NewsID
rs.open sql,conn,1,1
if rs.eof then
			%>
			<script language=javascript>
			history.back()
			alert("��Ҫ�鿴�����ݲ����ڻ��Ѿ�������Ա���������������ϵ��")
			</script>
			<%
		Response.End 
else	
	BigClassName=rs(1)
	SmallClassName=rs(2)
	SpecialName=rs(3)
	Title=rs(4)
	titleurl=rs(5)
	username=rs(6)
	Author=rs(8)
	if Author<>"" then AuthorR="&nbsp;���ߣ�"
	Original=rs(9)
	if Original<>"" then OriginalR="&nbsp;��Դ��"
	image=rs(11)
	UpdateTime=rs(12)
	content=rs(13)
	about=rs(14)
	click=rs(15)
	hot=rs(16)
end if
rs.Close


dim ReaderLevel,UserLevel
sql="select ReaderLevel from SmallClass where BigClassName='"&BigClassName&"' and SmallClassName='"&SmallClassName&"'"
rs.open sql,conn,1,1
ReaderLevel=rs(0) 
rs.close

if ReaderLevel<>0 then
if readfree=0 then  '���������ǻ�Ա���
if (isnull(session("xpUser")) or isnull(session("xpPwd")) or session("xpUser")="" or session("xpPwd")="") and readFree=0 then
	conn.close 
	set conn=nothing
	Response.Redirect "userlogin.asp"
	Response.End 
end if

if not(isnull(session("xpUser")) or isnull(session("xpPwd")) or session("xpUser")="" or session("xpPwd")="") then
	sql="select UserLevel,lockuser,LimitPoint,readnews,UserPoint from Users where UserName='"&Session("xpUser")&"' and PassWord='"&Session("xpPwd")&"'"
	rs.Open sql,conn,1,3
	if rs.eof then
			%>
			<script language=javascript>
			history.back()
			alert("�Բ��𣬻�Ա�ʺŻ�������������µ�½��")
			</script>
			<%
		Response.End 
	else
		dim rs1,theLimitPoint
		set rs1=conn.execute("SELECT LimitPoint FROM UserGrade where id="&rs(0)&"")
		theLimitPoint=rs1(0)
		rs1.close:set rs1=nothing	
		if rs(0)<ReaderLevel and rs(0)<7 then
			%>
			<script language=javascript>
			history.back()
			alert("��ĵȼ��ﲻ������Ŀ��Ҫ���<%=ReaderLevel%>���������Ŭ����")
			</script>
			<%
			Response.End 
		elseif rs(1)=1 then
			%>
			<script language=javascript>
			history.back()
			alert("����ʺű�����������ϵ����Ա��")
			</script>
			<%
			Response.End 
		elseif rs(2)>=theLimitPoint then
			%>
			<script language=javascript>
			history.back()
			alert("����Ķ������ѳ�������<%=theLimitPoint%>�Σ�����ߵȼ��Ի�ȡ���������")
			</script>
			<%
			Response.End 				
		else
			rs(3)=rs(3)+1
			rs(4)=rs(4)+1
			rs(2)=rs(2)+1
			if rs(0)<7 and int(rs(4))=int(point(rs(0)+1)) then rs(0)=rs(0)+1
			rs.Update
		end if
	end if
	rs.close
end if
end if
end if

conn.execute("update News Set Click=click+1 where NewsID=" & NewsID )

if titleurl="" or isnull(titleurl) then

sql="select * from News where username='"&username&"'"
rs.open sql,conn,1,1
dim InputCount
InputCount=rs.RecordCount
rs.Close
%>

<SCRIPT language=JavaScript>
	//����
	var currentpos,timer;	
	function initialize()
	{timer=setInterval("scrollwindow()",50);}
	function sc()
	{clearInterval(timer);}
	function scrollwindow()
	{currentpos=document.body.scrollTop;
	window.scroll(0,++currentpos);
	if (currentpos != document.body.scrollTop)
	sc();}
	document.onmousedown=sc
	document.ondblclick=initialize
</script>

<!--#include file="top.asp"-->
<table border="0" style="border-collapse: collapse" width="<%=TableWidth%>" cellspacing="0" cellpadding="0" bgcolor="<%=MainBColor%>" height="338">
<tr>
<%=OutTable("left")%>
<td align="left" valign="top" bgcolor="<%=CenterCColor%>"> <table width="100%" border="0" cellspacing="6" cellpadding="0">
<tr>
<td align="center" class=READNEWSTITLE>
<p>&nbsp;</p>
<p><b><%=title%></b></p>
</td>
</tr>
</table>
      <table width="100%" border="0" cellspacing="6" cellpadding="0">
        <tr> 
          <td align="center" class=NEWSREADME>
<hr size="1" noshade>
            ����ʱ�䣺<%=updateTime%><%=originalR%><%=original%><%=authorR%><%=author%>&nbsp;&nbsp;��� 
            <%=click%> ��<br>
            <br>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="6" cellpadding="0">
        <tr> 
          <td class=news><%=HtmlSelfEnCode(content,image)%></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="6" cellpadding="0">
        <tr> 
          <td align="right"><p>&nbsp;</p>
            <p>¼�룺<%=username%>&nbsp;[�� <%=InputCount%> ƪ]&nbsp;&nbsp;&nbsp;</p></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="6" cellpadding="0">
        <tr> 
          <td>&nbsp;��һƪ��
            <%movenews("-")%>
            <br> &nbsp;��һƪ��
            <%movenews("+")%>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="6" cellpadding="0">
        <tr> 
          <td width="50%"> 
            <%aboutnews%>
          </td>
          <td width="50%"> 
            <%thisspecial%>
          </td>
        </tr>
      </table>
<%
if ShowGBook=1 then
dim cols
if ShowNewsModelRight=1 then
cols=86
else
cols=111
end if
%>
<script language=javascript src="gbookshow.asp?NewsID=<%=NewsID%>&cols=<%=cols%>"></script>
<%
end if
%>
</td>
<%
if ShowNewsModelRight=1 then	'������ʼ
	Response.Write InTable("right")
	%>
	<td width="160" align="right" valign="top" bgcolor="<%=RightBgColor%>">
	<script language=javascript src="goodnews.asp?SmallClassName=<%=SmallClassName%>"></script>
	<script language=javascript src="hotimg.asp?SmallClassName=<%=SmallClassName%>"></script>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
	<tr>
	<td bgcolor="<%=RightCColor%>" background="<%=RightCImg%>" height="100%">&nbsp;</td>
	</tr>
	</table>
	<%	
end if	'��������
Response.Write OutTable("right")
%>
</tr>
</table>
<%
else
set rs=nothing
conn.close
set conn=nothing
Response.Redirect titleurl
end if

function thisspecial
	Response.Write "<table border='0' cellspacing='0' cellpadding='0' width='100%' style=""BORDER-LEFT: "&CenterTColor&" 1px double; BORDER-RIGHT: "&CenterTColor&" 1px double; BORDER-BOTTOM: "&CenterTColor&" 1px double""><tr><TD bgcolor="""&CenterBColor&""" background="""&CenterBImg&""" HEIGHT=1></TD></tr><tr><td width='100%' bgcolor='"&CenterTColor&"' background="""&CenterTImg&""" height='18' Class=maintitle>&nbsp;����ר�⣺"
	if specialname<>"��" then Response.Write "<a Class=maintitle href='Special.asp?SpecialName="& specialname &"' target=_self>" & specialname &"</a>"
	Response.Write "</td></tr><tr><TD bgcolor="""&CenterBColor&""" background="""&CenterBImg&""" HEIGHT=1></TD></tr><tr><td bgcolor='"&CenterCColor&"' background="""&CenterCImg&"""><table width='100%' border='0' cellspacing='5' cellpadding='0'><tr height=114><td width='100%' valign=top>"		
	sql="select top 5 "& NoContent &" from News where (checked=1 and NewsID<>" & NewsID & " and SpecialName='" & SpecialName & "') order by NewsID DESC"
	rs.open sql,conn,1,1
	if SpecialName="��" or rs.EOF or rs.bof then
		Response.Write "&nbsp;������Ϣ"
	else
		while not rs.EOF
			Response.Write shownewf
			if ShowNewsModelRight=1 then 
				Response.Write showTitle("MainContentS",30)
			else
				Response.Write showTitle("MainContentS",44)
			end if
			Response.Write showImg
			Response.Write showTime
			rs.MoveNext
		wend
		Response.Write "<div align='right'><a Class='MainMore' href='Special.asp?SpecialName="& specialname &"' target=_self>>>����</a>&nbsp;</div>"
	end if
	Response.write "</td></tr></table></td></tr></table>"
	rs.Close
end function

function aboutnews
	sql="select top 5 "&NoContent&" from news where checked=1 and about like '%" & about & "%' and title not like '" & title & "' order by newsid desc"
	rs.open sql,conn,1,1
	Response.Write "<table border=0 cellspacing=0 cellpadding=0 width=""100%"" style=""BORDER-LEFT: "&CenterTColor&" 1px double; BORDER-RIGHT: "&CenterTColor&" 1px double; BORDER-BOTTOM: "&CenterTColor&" 1px double""><tr><TD bgcolor="""&CenterBColor&""" background="""&CenterBImg&""" HEIGHT=1></TD></tr><tr><td width='100%' bgcolor="""&CenterTColor&"""  background="""&CenterTImg&""" height='18' class=maintitle>&nbsp;�����Ϣ��"&about&"</td></tr><tr><TD bgcolor="""&CenterBColor&""" background="""&CenterBImg&""" HEIGHT=1></TD></tr></tr><tr><td bgcolor="""&CenterCColor&""" background="""&CenterCImg&"""><table width='100%' border='0' cellspacing='5' cellpadding='0'><tr height=114><td valign=top>"
	if not rs.EOF and about<>"" then
		do while not rs.eof	
			Response.Write shownewf
			if ShowNewsModelRight=1 then 
				Response.Write showTitle("MainContentS",30)
			else
				Response.Write showTitle("MainContentS",44)
			end if
		Response.write showImg
		Response.write showTime
		rs.movenext
		loop
		Response.Write "<div align='right'><a Class='MainMore' href='showsearch.asp?keyword=" & about &"' target=_self>>>����</a>&nbsp;</div>"
	else
		Response.Write "&nbsp;������Ϣ"
	end if
	Response.Write "</td></tr></table></td></tr></table>"	
	rs.close
end function

function movenews(strmove)
	sql="select "& NoContent &" from News where checked=1 and NewsID=" & NewsID &strmove&"1"
	rs.open sql,conn,1,1
	if not rs.EOF then
		Response.write showTitle("MainContentS",70)
		Response.write showImg
		Response.write showTime
	else
		Response.write "�Ѿ�û����"
	end if
	rs.close
end function

set rs=nothing%>
<!--#include file="copyright.asp"-->