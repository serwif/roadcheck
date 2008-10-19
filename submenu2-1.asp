<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<HTML><HEAD>
<meta name="ProgId" content="FrontPage.Editor.Document">
<LINK href="./images/main.css" rel=stylesheet type=text/css>
<style type="text/css">
TD         { COLOR: #000000;FONT-SIZE: 12px }
.top       {MARGIN: 0px; }
.outdiv    { BORDER-BOTTOM: buttonhighlight 0px solid; BORDER-LEFT: buttonshadow 0px solid; BORDER-RIGHT: buttonhighlight 0px solid; BORDER-TOP: buttonshadow 1px solid; WIDTH: 100%}
.indiv1    { BORDER-BOTTOM: buttonshadow 1px solid; BORDER-LEFT: buttonshadow 0px solid; BORDER-RIGHT: buttonhighlight 0px solid; BORDER-TOP: buttonhighlight 1px solid; WIDTH: 100%}
.indiv     { BORDER-BOTTOM: buttonshadow 0px solid; BORDER-LEFT: buttonhighlight 0px solid; BORDER-RIGHT: buttonshadow 0px solid; BORDER-TOP: buttonhighlight 1px solid }
.handbtn   { BORDER-BOTTOM: buttonshadow 2px solid; BORDER-LEFT: buttonhighlight 2px solid; BORDER-RIGHT: buttonshadow 1px solid; BORDER-TOP: buttonhighlight 2px solid; WIDTH: 3px}
.sepbtn1   { BORDER-LEFT: buttonshadow 1px solid; BORDER-RIGHT: buttonhighlight 1px solid; WIDTH: 2px }

.showboder    { BORDER-BOTTOM: buttonshadow 1px solid; BORDER-LEFT: buttonhighlight 1px solid; BORDER-RIGHT: buttonshadow 1px solid; BORDER-TOP: buttonhighlight 1px solid }
.showboder1    {  BORDER-BOTTOM: buttonshadow 2px solid; BORDER-LEFT: buttonhighlight 2px solid; BORDER-RIGHT: buttonshadow 2px solid; BORDER-TOP: buttonhighlight 1px solid }
.showboder2    { BORDER-BOTTOM: buttonhighlight 1px solid; BORDER-LEFT: buttonshadow 2px solid; BORDER-RIGHT: buttonhighlight 1px solid; BORDER-TOP: buttonshadow 2px solid }
</style>
<SCRIPT language=JavaScript>
<!--
function over1() {title.className = "iconover";}
function out1() {title.className = "icon";}
function down1() {title.className = "icondown";}


function KeyDown(){
if(event.keyCode==9 || event.keyCode==13)        //屏蔽,
  {
   event.keyCode=0;
   event.returnValue=false;
   }
}
function inputmenu(menu){
	top.document.all.indexmenu.innerHTML=menu;
}
function window_go(url){
	top.main.location.href=url;
}
function over2(id) {id.className = "showboder";}
function out2(id) {id.className = "";}
function down2(id) {id.className = "showboder2";}

function movstar(a,time){
movx=setInterval("mov("+a+")",time)
}
function movover(){
clearInterval(movx)
}
function mov(a){
scrollx=new_date.document.body.scrollLeft
scrolly=new_date.document.body.scrollTop
scrolly=scrolly+a
new_date.window.scroll(scrollx,scrolly)
}
//-->
</SCRIPT>

<body bgcolor="#d6d3ce" leftmargin="0" topmargin="0" class=menu onkeydown="KeyDown()" onload="windowOnload()">
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

<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="122" align="center">
  <tr><td colspan="2" class=icon id=title width="122" height=23 onmouseover=over1() onmouseout=out1() background="./images/folder.gif" align=center>
    上报
  </td></tr>
  <tr align="center">
    <td align="center"  width="100" rowspan="2" height="100%">
      <iframe frameborder=0 framespacing=0 height="100%" marginheight=0 marginwidth=0 name=new_date scrolling=no 
        src="submenu2-1s.asp" width="100%" vspalc="0">
      </iframe>
    </td>
    <td align="center" width="22" valign="top" background="./images/folderbk2.gif">
      <table border="0" cellspacing="0" cellpadding="0" height="16">
        <tr><td height="10"></td></tr>
        <tr> 
          <td align="center" id="img1" onmouseover="over2(img1);" onmouseout="out2(img1);" onmousedown="down2(img1);" onmouseup="over2(img1);" width="16" height="16">
            <img name="imageup" style="cursor:hand" onMouseDown=movover();movstar(-10,1); 
              onMouseOut=movover(); onMouseOver=movstar(-4,20); onMouseUp=movover();movstar(-4,20); 
              src="images/arrowup.gif" alt="向上滚" width="16" height="16">
          </td>
        </tr>
      </table>
	</td>
  </tr>
  <tr align="center">
    <td align="center"  width="20" valign="bottom" background="./images/folderbk2.gif">
      <table border="0" cellspacing="0" cellpadding="0" height="16">
        <tr>
          <td align="center"  id="img2" onmouseover="over2(img2);" onmouseout="out2(img2);" onmousedown="down2(img2);" onmouseup="over2(img2);" width="16" height="16">
            <img name="imagedown" style="cursor:hand" onMouseDown=movover();movstar(10,1); 
              onMouseOut=movover(); onMouseOver=movstar(4,20); onMouseUp=movover();movstar(4,20); 
              src="images/arrowdown.gif" alt="向下滚" width="16" height="16">
          </td>
        </tr>
        <tr><td height="10"></td></tr>
      </table>
    </td>
  </tr>
</TABLE>
   
<SCRIPT language=VBScript>           
sub document_onclick()           
	window.Parent.frm1.rows =  "*,23"
end sub	           
           
sub title_OnMouseOver()           
	title.style.cursor = "hand"           
end sub	           
</SCRIPT>   
	<Script language="javascript">
	  function windowOnload(){
		  if (screen.availHeight<740){
				document.imageup.style.display=""
				document.imagedown.style.display=""
		  }
	  }
</script>

</BODY>
</HTML>