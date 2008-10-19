<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<!--#include file="./fcommon.asp"-->



<!--URL(/images/xmcdc.gif)-->
<html>
<head>
<title>Not Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="/main.css">
<style type=text/css>body{background-image: URL();background-position:bottom left;background-repeat: no-repeat;background-attachment:fixed;}</style>
</head>
<body leftmargin="0" topmargin="0">
<%noRightClick()%>
<br>
  <table width="530" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>
      <%if isempty(request("title")) then%>
        错误信息
      <%else%>
        <%=request("title")%>
      <%end if%>
      </b></td>
    </tr>
    <tr><td align="center"><br>
      <table width=90% align=center><tr><td>
      你在使用本系统时发生了错误，可能的原因为：<br>
      1、你没有登录或已经超时，请重新登录；<br>
      2、你没有权限使用相应的功能，请向系统管理员确认你的权限。
      </td></tr></table>
      <br>
    </td></tr>
  </table>
</body>
</html>
