<%
Response.Buffer = True
Response.Expires = 0
Response.CacheControl = "no-cache"

session.timeout=20
dim where
select case PageName
case "default" 
 where=""
case "BigClass" 
 where=request_BigClassName
case "SmallClass" 
 where="<a class=top4 href=""BigClass.asp?BName="& request_BigClassName & "&BType=" & request_BigClassType & """>" & request_BigClassName &"</a> > "&request_SmallClassName
case "MoreAnnounce" 
 where="���๫��"
case "Search" 
 where="����"
case "shownews" 
 where="�Ķ�����"
case "gbook" 
 where="����"
case "SpecialList" 
 where="ר���б�"
case "Special" 
 where="<a class=top4 href=Special.asp>ר��</a> > "&request_SpecialName
case "SpecialList" 
 where="ר���б�"
case "ImageList" 
 where="ͼƬ�б�"
case "hottxt" 
 where="�ȵ��б�"
case "goodnews" 
 where="�Ƽ��б�"
case "focusnews" 
 where="�����б�"   
case "UserReg" 
 where="ע���Ա��һ��"
case "UserReg2" 
 where="ע���Ա�ڶ���"
case "UserRegPost" 
 where="�ɹ�ע���Ա"
case "UserLogin" 
 where="��Ա��½"
case "UserList" 
 where="��ʱ����ʾ�û��б�"
case "UserList2" 
 where="��������ʾ�û��б�"
case "UserListgirl" 
 where="��ʾŮͬ���û��б�"
case "UserListboy" 
 where="��ʾ��ͬ���û��б�"
case "UserModify" 
 where="��Ա�޸�����"
case "UserSave" 
 where="��Ա�޸����ϳɹ�"
case "user_NewsAdd" 
 where="��Ա�������" 
case "admin_login" 
 where="����Ա��½"  
case else
 where="δ֪"
end select
%>
<html>
<head>
<title><%=WebTitle%></title>
<meta http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel="stylesheet" type="text/css" href="<%=weburl%>style.css">
</head>
<body topmargin=0 leftmargin=0 bgcolor="<%=BgColor%>" background="<%=weburl&bgImg%>">
<div align=center>
<!--<table style="BORDER-TOP: <%=TopBColor%> 1px double;BORDER-LEFT: <%=TopBColor%> 1px double;BORDER-RIGHT: <%=TopBColor%> 1px double;" border=0 width=<%=TableWidth%> cellspacing=0 cellpadding=0 bgcolor=<%=Top1bgColor%>>
	<tr>
	<td width=25><img src="<%=weburl%>images/TOPL.gif" hspace=3 border=0></td>
	<td><a class=top1 href="default.asp"><%=WebTitle%></a></td>
	<td align=right class=top1> 
      <%if ShowUserLogin=1 then
	  if writeOpen=1 then%>
      | <a class=top1 href="user_newsadd1.asp" target="_blank">��Ҫ����</a> 
	  <%end if%>
	  | <a class=top1 href=<%=weburl%>user/UserLogin.asp>��Ա��½</a> 
      | <a class=top1 href=<%=weburl%>user/UserLogOut.asp>��Ա�˳�</a> | <a class=top1 href=<%=weburl%>user/UserReg.asp>��Աע��</a> 
      | <a class=top1 href=<%=weburl%>user/UserList.asp>��Ա�б�</a> | 
      <%end if%>
    </td>
	<td width=25 align=right><img src="<%=weburl%>images/TOPR.gif" width=15 height=15 hspace=3 border=0></td>
	</tr>
</table>	
	
<table style="BORDER-TOP: <%=TopBColor%> 1px double;BORDER-LEFT: <%=TopBColor%> 1px double;BORDER-RIGHT: <%=TopBColor%> 1px double" border=0 width=<%=TableWidth%> cellspacing=0 cellpadding=0 bgcolor="<%=Top2bgColor%>">
 	<tr><td valign=middle width=100% ><script language="javascript" src="<%=weburl%>Ads.asp"></script></td></tr>
</table>-->
	
<script language="javascript" src="<%=weburl%>topmenu.asp"></script>
	
<table width=<%=TableWidth%> border=0 cellspacing=0 cellpadding=0 BGCOLOR=<%=Out2Color%>>
	<tr>
	<TD BGCOLOR=<%=Out1Color%> WIDTH=1></TD>
	<TD WIDTH=<%=out2width%> BGCOLOR=<%=Out2Color%>>
	<td style="BORDER-BOTTOM: <%=Out3Color%> 1px double" width=20 align=center><img src=<%=weburl%>images/where.gif border=0></td><td style="BORDER-BOTTOM: <%=Out3Color%> 1px double" height=20 align=left class=top4>��ǰλ�ã�<a class=top4 href=<%=weburl%>default.asp>��ҳ</a><%if pagename<>"default" then%>&nbsp;>&nbsp;<%end if%><%=where%></td>		
	<!--<td style="BORDER-BOTTOM: <%=out3color%> 1px double" height=20 align=right>
	<a href="<%=weburl%>SpecialList.asp" class=top4>ר��</a>&nbsp;
	<a href="<%=weburl%>HottxtList.asp" class=top4>�ȵ�</a>&nbsp;
	<a href="<%=weburl%>FocusNewsList.asp" class=top4>����</a>&nbsp;
	<a href="<%=weburl%>GoodNewsList.asp" class=top4>�Ƽ�</a>&nbsp;
	<a href="<%=weburl%>ImageList.asp" class=top4>ͼƬ</a>&nbsp;
	<%if PageName="shownews" then%>
		</td>
    <td WIDTH=24 style="BORDER-BOTTOM: <%=out3color%> 1px double" align="center"><a href="javascript:open_window('<%=weburl%>user/user_sendnews.asp?newsid=<%=newsid%>&title=<%=title%>','sendnews','width=450,height=500')"><img src="<%=weburl%>images/email.gif" alt="���ͱ�����" border="0"></a></td>
    <td WIDTH=24 style="BORDER-BOTTOM: <%=out3color%> 1px double" align="center" valign="bottom"><a href="javascript:window.print()"><img src="<%=weburl%>images/printer.gif" width=16 height=14 border=0 alt="��ӡ��ҳ"></a> 
      <%end if%>
    </td>
	<TD WIDTH=<%=out2width%> BGCOLOR=<%=Out2Color%>></TD>
	<TD BGCOLOR=<%=Out1Color%> WIDTH=1></TD>	-->
	</tr>
</table>