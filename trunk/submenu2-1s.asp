<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312">
<title>上报菜单</title>
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
  <a href="addfmc.asp" target=right><img src="./images/235.gif" border=0><br>花名册</a><br><br>
  <a href="addbgk.asp?tjbb=0100000000" target=right><img src="./images/235.gif" border=0><br>债务情况</a><br><br>
  <a href="addbgk.asp?tjbb=0200000000" target=right><img src="./images/235.gif" border=0><br>人员及经费</a><br><br>
  <a href="addbgk.asp?tjbb=0300000000" target=right><img src="./images/235.gif" border=0><br>通行费收入</a><br><br>
  <a href="addbgk.asp?tjbb=0400000000" target=right><img src="./images/235.gif" border=0><br>稽查业务</a><br><br>
  <a href="addwz.asp" target=right><img src="./images/235.gif" border=0><br>文章管理</a><br><br>
</td></tr>
</table>

</BODY>
</HTML>
