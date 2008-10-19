<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312">
<title>用户管理菜单</title>
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

<TABLE cellSpacing=0 cellPadding=0 border=0 width="100" align=center>
<tr><td align=center><br>
  <a href="muser.asp" target=right><img src="./images/200.gif" border=0><br>用户管理</a><br><br>
  <a href="mrizhi.asp" target=right><img src="./images/200.gif" border=0><br>日志管理</a><br><br>
</td></tr>
</table>

</BODY>
</HTML>
