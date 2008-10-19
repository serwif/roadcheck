<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312">
<title>个人设置菜单</title>
<link rel="stylesheet" type="text/css" href="./main.css">
</head>
<body background="./images/folderbk1.gif" leftmargin="0" topmargin="0">
<script language=JavaScript>
<!--
var message="";
///////////////////////////////////
function clickIE() {if (document.all) {(message);return false;}}
function clickNS(e) {if 
(document.layers||(document.getElementById&&!document.all)) {
if (e.which==2||e.which==3) {(message);return false;}}}
if (document.layers) 
{document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;}
else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;}
document.oncontextmenu=new Function("return false")
// --> 
</script>
<!--href="chooseskin.asp" -->

<TABLE cellSpacing=0 cellPadding=0 border=0 width="100" align=center>
<tr><td align=center><br>
  <!--<a target=right><img src="./images/237.gif" border=0><br>系统皮肤选择</a><br><br>-->
  <a href="changepwd.asp" target=right><img src="./images/236.gif" border=0><br>改变密码</a><br><br>
  <!--<a href="writeGzjh.asp" target=right><img src="./images/123.gif" border=0><br>掇写工作计划</a><br><br>-->
  <!--<a href="LookGzjh.asp" target=right><img src="./images/123.gif" border=0><br>查看工作计划</a><br><br>-->
  <!--<a href="writeim.asp" target=right><img src="./images/123.gif" border=0><br>发送即时信息</a><br><br>-->
  <!--<a href="askclearim.asp" target=right><img src="./images/144.gif" border=0><br>清除即时信息</a><br><br>-->

</td></tr>
</table>

</BODY>
</HTML>
