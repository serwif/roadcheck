<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<html>
<head>
<title>错误信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="/main.css">
</head>
<body leftmargin="0" topmargin="0">
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

  <table width=100% height=100% border=0 align=center><tr><td valign="middle">
  <table background="/images/errorbk<%=trim(cstr(int(rnd*2+1)))%>.jpg" width=100% height="80" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr><td>
    <table width="300" border="0" cellspacing="0" cellpadding="0" align=right>
      <tr><td>
        <%if request("errid")="1" then%>
          无法连接数据库！
        <%end if%>
      </td></tr>
    </table>
  </td></tr>
  </table>
  <table width="100%">
    <tr>
      <td><a href="login.asp"><font color=696969>返回首页</font></a></td>
      <td align="right"><font color=696969>三明市交通局--三明聚隆网络有限责任公司</font></td>
    </tr>
  </table>
  </td></tr></table>

</body>
</html>
