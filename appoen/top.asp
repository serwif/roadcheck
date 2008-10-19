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
 where="更多公告"
case "Search" 
 where="搜索"
case "shownews" 
 where="阅读文章"
case "gbook" 
 where="评论"
case "SpecialList" 
 where="专题列表"
case "Special" 
 where="<a class=top4 href=Special.asp>专题</a> > "&request_SpecialName
case "SpecialList" 
 where="专辑列表"
case "ImageList" 
 where="图片列表"
case "hottxt" 
 where="热点列表"
case "goodnews" 
 where="推荐列表"
case "focusnews" 
 where="焦点列表"   
case "UserReg" 
 where="注册会员第一步"
case "UserReg2" 
 where="注册会员第二步"
case "UserRegPost" 
 where="成功注册会员"
case "UserLogin" 
 where="会员登陆"
case "UserList" 
 where="按时间显示用户列表"
case "UserList2" 
 where="按积分显示用户列表"
case "UserListgirl" 
 where="显示女同胞用户列表"
case "UserListboy" 
 where="显示男同胞用户列表"
case "UserModify" 
 where="会员修改资料"
case "UserSave" 
 where="会员修改资料成功"
case "user_NewsAdd" 
 where="会员添加文章" 
case "admin_login" 
 where="管理员登陆"  
case else
 where="未知"
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
      | <a class=top1 href="user_newsadd1.asp" target="_blank">我要发表</a> 
	  <%end if%>
	  | <a class=top1 href=<%=weburl%>user/UserLogin.asp>会员登陆</a> 
      | <a class=top1 href=<%=weburl%>user/UserLogOut.asp>会员退出</a> | <a class=top1 href=<%=weburl%>user/UserReg.asp>会员注册</a> 
      | <a class=top1 href=<%=weburl%>user/UserList.asp>会员列表</a> | 
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
	<td style="BORDER-BOTTOM: <%=Out3Color%> 1px double" width=20 align=center><img src=<%=weburl%>images/where.gif border=0></td><td style="BORDER-BOTTOM: <%=Out3Color%> 1px double" height=20 align=left class=top4>当前位置：<a class=top4 href=<%=weburl%>default.asp>首页</a><%if pagename<>"default" then%>&nbsp;>&nbsp;<%end if%><%=where%></td>		
	<!--<td style="BORDER-BOTTOM: <%=out3color%> 1px double" height=20 align=right>
	<a href="<%=weburl%>SpecialList.asp" class=top4>专题</a>&nbsp;
	<a href="<%=weburl%>HottxtList.asp" class=top4>热点</a>&nbsp;
	<a href="<%=weburl%>FocusNewsList.asp" class=top4>焦点</a>&nbsp;
	<a href="<%=weburl%>GoodNewsList.asp" class=top4>推荐</a>&nbsp;
	<a href="<%=weburl%>ImageList.asp" class=top4>图片</a>&nbsp;
	<%if PageName="shownews" then%>
		</td>
    <td WIDTH=24 style="BORDER-BOTTOM: <%=out3color%> 1px double" align="center"><a href="javascript:open_window('<%=weburl%>user/user_sendnews.asp?newsid=<%=newsid%>&title=<%=title%>','sendnews','width=450,height=500')"><img src="<%=weburl%>images/email.gif" alt="发送本文章" border="0"></a></td>
    <td WIDTH=24 style="BORDER-BOTTOM: <%=out3color%> 1px double" align="center" valign="bottom"><a href="javascript:window.print()"><img src="<%=weburl%>images/printer.gif" width=16 height=14 border=0 alt="打印本页"></a> 
      <%end if%>
    </td>
	<TD WIDTH=<%=out2width%> BGCOLOR=<%=Out2Color%>></TD>
	<TD BGCOLOR=<%=Out1Color%> WIDTH=1></TD>	-->
	</tr>
</table>