<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if session("username")=""  or (instr(session("power"),",1,")=0 and instr(session("power"),",2,")=0) then
  Response.Redirect("notlogin.asp")
end if
%>

<!--#include file="fcommon.asp"-->
<!--#include file="dtp.asp"-->
<!--#include FILE="upload_5xSoft.inc"-->
<%
dim conn, mode, username, rs, sql,rs1,rsMX, errmsg, founderror, s, t, i, fl, dq,odq, cpage,kpbm,st
dim unit_code,unit_name,ajlb_code,fxlb_code,afsj,ajjs,czshj,explain,sday,ajbh,zbzcy,zp,xm

if not isempty(request("mode")) then
    mode = clng(request("mode"))
else
    mode=1
end if
if not isempty(request("username")) then
    username = request("username")
else
    username = ""
end if
if not isempty(request("unit_code")) then
    unit_code = request("unit_code")
else
    unit_code =""
end if
if not isempty(request("xm")) then
    xm = request("xm")
else
    xm = ""
end if

sub opendb()
  set conn=server.createobject("ADODB.CONNECTION")
  conn.open sysconstr
end sub

sub closedb()
  conn.Close
  set conn=nothing
end sub

sub showchead()
%>
  <html>
  <head>
  <meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312">
  <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
  <title>花名册</title>
  <link rel="stylesheet" type="text/css" href="./main.css">
  </head>
<script language="javascript">
  <!--
  function surfto(list){
   var myindex1=list.selectedIndex;
   if (myindex1!=0 & myindex1!=1){ location.href=list.options[list.selectedIndex].value }
  }
  function goto(list){
   location.href=list.options[list.selectedIndex].value
  }
  //-->
function trim(word) {
  while(word.length>0) {
    if(word.substring(0,1)==" ")
       word=word.substring(1,word.length)
    else
     if(word.substring(0,2)=="  ")
        word=word.substring(2,word.length)
     else
       break
}
   while(word.length>0) {
     if(word.substring(word.length-1,word.length)==".")
        word = word.substring(0,word.length-1)
     else
        if(word.substring(word.length-2,word.length)=="  ")
           word = word.substring(0,word.length-2)
        else
          break
 }
return word
}

function isnumber(word){
  var i=0;
  var result=true;
  for(i=0;i<word.length;i++){
    if(word.charAt(i)<'0'||word.charAt(i)>'9'){
	   result=false;
	   break;
  }
 }
return result;
} 

function check()
{
   location.href="addfmc.asp?mode=1&DisDate=8&sday=" + document.all.afsj.value+"&unit_code="+document.form1.unit_code.value; 
   return false;   
}

function check_form() 
{ 
  if(trim(document.form1.xm.value)==""){
   alert("请完整填入姓名!"); 
   return false; 
   }
  return true;
} 

  function Getseconditem(i,j)
  {//求大类的小类列表
   var unit_code;
   if(j==1)
     unit_code=document.input1.unit_code1.options[document.input1.unit_code1.selectedIndex].value;
   else
     {if(j==2)
        unit_code=document.input1.unit_code2.options[document.input1.unit_code2.selectedIndex].value; 
      else
        {if(j==3)
           unit_code=document.input1.unit_code3.options[document.input1.unit_code3.selectedIndex].value; 
        } 
     }
   //alert(i);
   if(i==102)
     location.href="addfmc.asp?mode=2&unit_code="+unit_code;
   else
     location.href="addfmc.asp?mode=3&unit_code="+unit_code;
   return false;
  }

</script>  
  <body>
  <%noRightClick()%>
  <!--<table width="90%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28"><td align="center">
      <b>花名册</b>
    </td></tr>
  </table>-->
<%
end sub

sub showctail()
%>
  </body>
  </html>
<%
end sub

sub ShowInputForm1(mode,errmsg)
  'on error resume next
  showchead()
  if mode = 2 then
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn, 1, 1%>
    <form name="form1" method="post" onsubmit="return check_form()" action="addfmc.asp?mode=2">
  <%else
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn, 1, 1
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from fmc where bh='" + request("odq") + "'", conn, 1, 1
    %>
    <form name="form1" method="post" onsubmit="return check_form()" action="addfmc.asp?mode=3&odq=<%=request("odq")%>">
  <%end if%>
  <table width="530" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>花名册</b></td>
    </tr>
    <tr>
      <td align=center>
        <table width=100%><tr>
        </table>
      </td>
    </tr>
    <tr>
      <td align="center"><br><font color=red>
        <%if Trim(ErrMsg) <> "" then%>
          <%=errmsg%>
        <%else%>
          请填写花名册内容。
          <% if mode=2 then%>
            <input name="odq1" type="hidden" value="">
          <% else %>
            <input name="odq1" type="hidden" value="<%=rs("bh")%>">
          <% end if %>
        <%end if%>
        </font><br><br>
      </td>
    </tr>
    <tr>
      <td align=center>
        <table width="100%" border="0" cellspacing="1" bgcolor="#cccccc">
          <!--DWLayoutTable-->
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">单位：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
  	        地&nbsp;&nbsp;区
            <%if mode=2 then%> 
              <select name="unit_code1" style="HEIGHT:17px;WIDTH:119px" disabled>
            <%else%>
              <select name="unit_code1" style="HEIGHT:17px;WIDTH:119px" disabled>
            <%end if%>
            <%while not rs1.EOF 
              if trim(unit_code)="" then unit_code=left(session("unit_code"),unit_len1)%>
              <option value="<%=trim(rs1("unit_code"))%>"<%if left(unit_code,unit_len1)=left(rs1("unit_code"),unit_len1) then %> selected <% end if %>><%=trim(rs1("unit_name"))%></option>
              <%rs1.MoveNext 
            WEND%>
            </select>
            <br>收费站
            <%if mode=2 then%>
              <%if right(session("unit_code"),unit_len0-unit_len2) = unit_str2 and mid(session("unit_code"),unit_len1+1,unit_len2-unit_len1)="00" then%>
                <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" >
              <%else%>
                <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" disabled>
              <%end if%>
            <%else%>
              <%if right(session("unit_code"),unit_len0-unit_len2) = unit_str2 and mid(session("unit_code"),unit_len1+1,unit_len2-unit_len1)="00" then%>
                <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" >
              <%else%>
                <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" disabled>
              <%end if%>
            <%end if%>
            <option value="" <%if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then %> selected <% end if %>></option>
            <%rsMX.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn, 1, 1
            while not rsMX.EOF
              %>
              <option value="<%=trim(rsMX("unit_code"))%>"<%if left(unit_code,unit_len2)=left(rsMX("unit_code"),unit_len2) then %> selected <% end if %>><%=trim(rsMX("unit_name"))%></option>
              <%rsMX.MoveNext 
            WEND
            rsMX.close%>
            </select>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">姓名：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="xm" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="xm" size="20" maxlength="20" value='<%=rs("xm")%>'>
                <input name="oxm" type="hidden" value="<%=rs("xm")%>">
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">照片：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="file" name="zp" maxlength="255">
              <%else%>
                <input type="file" name="zp" maxlength="255">
              <%end if%>
              请先将照片上传到管理系统目录
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">职务：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="zw" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="zw" size="20" maxlength="20" value='<%=rs("zw")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">工作：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="gz" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="gz" size="20" maxlength="20" value='<%=rs("gz")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee" align="right">工人/干部： </td>
            <td height="23" colspan="7" bgcolor="#FFFFFF">
            <%if mode=2 then%>
              <input type="radio" name="sfgrgb" value='工人' checked>工人
              <input type="radio" name="sfgrgb" value='干部'>干部
              <input type="radio" name="sfgrgb" value='合同工'>合同工
              <input type="radio" name="sfgrgb" value='借用'>借用
              <input type="radio" name="sfgrgb" value='临时工'>临时工
            <%else%>
              <input type="radio" name="sfgrgb" value='工人' <%if rs("sfgrgb")="工人" then %>checked<%end if%>>工人              
              <input type="radio" name="sfgrgb" value='干部' <%if rs("sfgrgb")="干部" then %>checked<%end if%>>干部
              <input type="radio" name="sfgrgb" value='合同工' <%if rs("sfgrgb")="合同工" then %>checked<%end if%>>合同工
              <input type="radio" name="sfgrgb" value='借用' <%if rs("sfgrgb")="借用" then %>checked<%end if%>>借用
              <input type="radio" name="sfgrgb" value='临时工' <%if rs("sfgrgb")="临时工" then %>checked<%end if%>>临时工
            <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">民族：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="mz" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="mz" size="20" maxlength="20" value='<%=rs("mz")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee" align="right">性别： </td>
            <td height="23" colspan="7" bgcolor="#FFFFFF">
            <%if mode=2 then%>
              <input type="radio" name="xb" value='男' checked>男
              <input type="radio" name="xb" value='女'>女
            <%else%>
              <input type="radio" name="xb" value='男' <%if rs("xb")="男" then %>checked<%end if%>>男              
              <input type="radio" name="xb" value='女' <%if rs("xb")="女" then %>checked<%end if%>>女
            <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">籍贯：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="jg" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="jg" size="20" maxlength="20" value='<%=rs("jg")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">出生年月：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="csly" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="csly" size="20" maxlength="20" value='<%=rs("csly")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">毕业时间：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="wfcdbysj" size="10" maxlength="12" readonly value=''>
              <%else%>
                <input type="text" name="wfcdbysj" size="10" maxlength="12" readonly value='<%=todatestr(rs("wfcdbysj"))%>'>
              <%end if%>
              <A onclick="show_cele_date(change1,'','',wfcdbysj)"><IMG align=top border=0 height=25 name=change1 src="images\calendar.gif" width=26></A>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">院校：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="wfcdyx" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="wfcdyx" size="20" maxlength="20" value='<%=rs("wfcdyx")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">学历：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="wfcdxl" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="wfcdxl" size="20" maxlength="20" value='<%=rs("wfcdxl")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">入伍年月：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="rwly" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="rwly" size="20" maxlength="20" value='<%=rs("rwly")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee" align="right">党/团员： </td>
            <td height="23" colspan="7" bgcolor="#FFFFFF">
            <%if mode=2 then%>
              <input type="radio" name="dty" value='无党派人士' checked>无党派人士
              <input type="radio" name="dty" value='党员'>党员
              <input type="radio" name="dty" value='团员'>团员
            <%else%>
              <input type="radio" name="dty" value='无党派人士' <%if rs("dty")="无党派人士" then %>checked<%end if%>>无党派人士             
              <input type="radio" name="dty" value='党员' <%if rs("dty")="党员" then %>checked<%end if%>>党员             
              <input type="radio" name="dty" value='团员' <%if rs("dty")="团员" then %>checked<%end if%>>团员
            <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">收费证号：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="sfzh" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="sfzh" size="20" maxlength="20" value='<%=rs("sfzh")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">执法证号：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="zfzh" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="zfzh" size="20" maxlength="20" value='<%=rs("zfzh")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">奖惩记录：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <textarea type="text" name="jc" size="20" cols=60 rows=2></textarea>
              <%else%>
                <textarea type="text" name="jc" size="20" cols=60 rows=2><%=rs("jc")%></textarea>
              <%end if%>
            </td>
          </tr>
        </table>
        <p> 
        <input class="buttonface" type="submit" name="Submit" value=" 提 交 ">
        &nbsp; 
        <INPUT class="buttonface" type=reset onclick="{if(confirm('该项操作要清除全部的内容，您确定要清除吗?')){return true;}return false;}" value=" 重 写 " id=reset1 name=reset1>
        </p>   
        </div>
      </td>
    </tr>
    </table>
  </form>
<%
  if mode = 2 then
    rs1.close
    set rs=nothing
    closedb()
  elseif mode = 3  then
    rs1.close
    rs.close
    set rs=nothing
    closedb()
  end if
  showctail
end sub

sub ShowInputForm3(ErrMsg)
  'on error resume next
  showchead()%>
  <form method="POST" action="addfmc.asp?mode=5&username=<%=username%>" name="input3">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="addfmc.asp?mode=8&username=<%=username%>">返回</a>]
      </td>
    </tr>
    <tr><td><hr size="1" noshade width="100%"></td></tr>
    <tr><td>
      <table width="450" border="0" cellspacing="1" cellpadding="0" align="center" bordercolor="#FFCC33">
        <tr>
        <%if Trim(ErrMsg) <> "" then%>
          <td align=center><%=errmsg%><br><br></td>
        <%else%>
          <td align=center>请输入查找条件<br><br></td>
        <%end if%>
        </tr>
        <tr>
          <td align=center><input type="text" name="dq" size="60" maxlength="20" class="smallInput" value="<%=request("dq")%>"></td>
        </tr>
        <tr align="center">
          <td><br><input class="buttonface" type="submit" value=" 开始查找 " id=submit1 name=submit1></td>
        </tr>
      </table>
    </td></tr>
    <tr><td><hr size="1" noshade width="100%"></td></tr>
  </table>
  </form>
<%
  showctail
end sub

if mode=1 then
  '显示
  if not isEmpty(request("page")) and isnumeric(request("page")) then
    cpage = clng(request("page"))
  else
    cpage = 1
  end if
  showchead()
  'Response.Write "<br>"
  opendb()
  set rs=server.createobject("adodb.recordset")
  set rs1=server.createobject("adodb.recordset")
  if right(session("unit_code"),unit_len0-unit_len2) = unit_str2 and mid(session("unit_code"),unit_len1+1,unit_len2-unit_len1)="00" then
	sql="select * from fmc where unit_code like '"+left(session("unit_code"),unit_len1)+"%' "
  else
    sql="select * from fmc where unit_code='"+session("unit_code")+"' "
  end if
  sql=sql+" order by xdw,xm"
  'response.write sql
  rs.open sql, conn, 1, 1
  %>
  <form name="form1">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td height="23"  align="left" bgcolor="#FFFFFF" >
        <input name="unit_code" type="hidden" value="<%=request("unit_code")%>">
      </td>
    </tr>
  </tabel>
  </form>
  <%if rs.recordcount <> 0 then
  rs.movefirst
  rs.CacheSize = 5
  rs.PageSize = 10
  if cpage>rs.pagecount then cpage=1
  rs.AbsolutePage = cpage%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
    </tr>
    <tr>
      <td valign="bottom">第<%=cstr(cpage)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个花名册记录</td>
      <td align="right">
        <%if instr(session("power"),",1,")=1 then %>
          [<a href="addfmc.asp?mode=2&unit_code=<%=session("unit_code")%>">添加</a>]
        <%end if %>
        <%if cpage <> 1 then%>
          [<a href="addfmc.asp?mode=1&page=<%=cstr(cpage-1)%>">上一页</a>]
        <%end if%>
        <%if cpage <> rs.PageCount then%>
          [<a href="addfmc.asp?mode=1&page=<%=cstr(cpage+1)%>">下一页</a>]
        <%end if%>
        <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage then%>
                <option selected value="addfmc.asp?mode=1&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="addfmc.asp?mode=1&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%end if%>
             <%next%>
          </select>
        <%end if%>
      </td>
    </tr>
    <tr>
      <td colspan="2">
        <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
          <tr bgcolor=<%=skincolor()%>>
            <td width=10% align=center>姓名</td>
            <td width=10% align=center>单位</td>
            <td width=10% align=center>收费证号</td>
            <td width=30% align=center>执法证号</td>
            <%if instr(session("power"),",2,")>0 then %>
              <td width=15% align=center>操作</td>
            <%end if%>
          </tr>
          <%fl = False
          for i = 1 to rs.PageSize
            if not rs.EOF then
              if fl then%>
                <tr bgcolor="#eeeeee">
              <%else%>
                <tr>
              <%end if%>
              <td align=center>
                <%
				if isnull(rs("xm")) then
				  response.write "&nbsp;"
				elseif rs("xm")="" then 
                  response.write "&nbsp;"
                else
                  response.write rs("xm")
                end if%>
              </td>
			  <td align=center>
                <%
				if isnull(rs("xdw")) then
				  response.write "&nbsp;"
				elseif rs("xdw")="" then 
                  response.write "&nbsp;"
                else
                  response.write rs("xdw")
                end if%>
              </td>
              <td align=center>
                <%
				if isnull(rs("sfzh")) then
				  response.write "&nbsp;"
				elseif rs("sfzh")="" then 
                  response.write "&nbsp;"
                else
                  response.write rs("sfzh")
                end if%>
              </td> 
			  <td align=center>
                <%
				if isnull(rs("zfzh")) then
				  response.write "&nbsp;"
				elseif rs("zfzh")="" then 
                  response.write "&nbsp;"
                else
                  response.write rs("zfzh")
                end if%>
              </td>
              <%if instr(session("power"),",2,")>0 then%>
                <td align=center>
                  <a href="addfmc.asp?mode=3&unit_code=<%=rs("unit_code")%>&odq=<%=trim(rs("bh"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="addfmc.asp?mode=4&unit_code=<%=rs("unit_code")%>&dq=<%=trim(rs("bh"))%>&xm=<%=trim(rs("xm"))%>"><img src="./images/del.gif" border=0></a>
                </td>
              <%end if%>
            </tr>
            <%rs.MoveNext
            fl = not fl
          end if
        next%>
      </table>
    </td>
  </tr>
</table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          <%if instr(session("power"),",1,")=1 then%>
          [<a href="addfmc.asp?mode=2&unit_code=<%=session("unit_code")%>">添加</a>]
          <%end if %>
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有花名册记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()
elseif mode=2 or mode=3 then
  '添加及修改
  if trim(request("xm")) <> "" then
    if trim(request("unit_code2"))="" then
      unit_code=request("unit_code1")
    else
      unit_code=request("unit_code2")
    end if
    if unit_code="" then
      unit_code=session("unit_code")
    end if
    if len(request("zp"))>255 then
      zp=left(request("zp"),255)
    else
      zp=request("zp")
    end if
    username=trim(session("username"))
    czshj=now()
    FoundError=false
    ErrMsg=""
    if not FoundError then
      if mode=2 then
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册，判断有无重复的同一时间做两件事
        rs.open "select bh from fmc where unit_code='"+unit_code+"' and xm='" + request("xm") + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "姓名重复"
          FoundError = True
        end if
        rs.close
        if trim(request("sfzh"))<>"" then
          rs.open "select bh from fmc where unit_code='"+unit_code+"' and sfzh='" + request("sfzh") + "'", conn, 1, 1
          if rs.recordcount<>0 then
            if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
            ErrMsg = ErrMsg + "收费证号重复"
            FoundError = True
          end if
          rs.close
        end if
        if trim(request("zfzh"))<>"" then
          rs.open "select bh from fmc where unit_code='"+unit_code+"' and zfzh='" + request("zfzh") + "'", conn, 1, 1
          if rs.recordcount<>0 then
            if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
            ErrMsg = ErrMsg + "执法证号重复"
            FoundError = True
          end if
          rs.close
        end if
        set rs=nothing
        closedb()
      else
        '看改过的案件编号是否存在
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册，判断有无重复的同一时间做两件事
        rs.open "select bh from fmc where unit_code='"+unit_code+"' and xm='" + request("xm") + "' and bh<>'"&request("odq") &"'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "姓名重复"
          FoundError = True
        end if
        rs.close
        if trim(request("sfzh"))<>"" then
          rs.open "select bh from fmc where unit_code='"+unit_code+"' and sfzh='" + request("sfzh") + "' and bh<>'"&request("odq") &"'", conn, 1, 1
          if rs.recordcount<>0 then
            if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
            ErrMsg = ErrMsg + "收费证号重复"
            FoundError = True
          end if
          rs.close
        end if
        if trim(request("zfzh"))<>"" then
          rs.open "select bh from fmc where unit_code='"+unit_code+"' and zfzh='" + request("zfzh") + "' and bh<>'"&request("odq") &"'", conn, 1, 1
          if rs.recordcount<>0 then
            if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
            ErrMsg = ErrMsg + "执法证号重复"
            FoundError = True
          end if
          rs.close
        end if
        set rs=nothing
        closedb()        
      end if
    end if
    if FoundError=true then
      ShowInputForm1 mode,errmsg
    else
	  '上传图片至指定目录(服务器软件发布目录下的PHOTO目录)
	  dim fso,CopyFile
	  dim upNum
      dim upload,file,formName,formPath,iCount,filename,fileExt
      if trim(zp)<>"" then
	    'Set fso = CreateObject("Scripting.FileSystemObject")
        'UploadPath =server.mappath("\photo\")' "\\wwwwSB\temp\"  '--这里的WWWWSB是SB的机器名temp为完全控制的目录'改为在fcommon.asp中定义
		'response.write UploadPath & "&nbsp;"
		'response.write zp  & "&nbsp;"
		'response.write getFileName(zp) & "&nbsp;"
        'Set CopyFile = fso.GetFile(zp)  '-->这里的test为用户上传到SA的目录
        'CopyFile.copy UploadPath
        'set fso = nothing
		'set upload=new upload_5xSoft
		'formPath=UploadPath
		'if right(formPath,1)<>"/" then formPath=formPath&"/"
		'file.SaveAs Server.mappath(FileName)
	  end if
	  
	  '保存数据
      if mode=2 then
	    '判断是否已经存在同一时间段同一个人的工作记录
	    opendb()
	    set rs=server.createobject("adodb.recordset")
        if FoundError=true then
	      set rs=nothing
	      closedb()
	      ShowInputForm1 mode,errmsg
	    else
          kpbm=right(cstr(year(date)),4)
	      if month(date)<10 then
	        kpbm=kpbm&"0"&cstr(month(date))
	      else
	        kpbm=kpbm&cstr(month(date))
	      end if
	      if day(date)<10 then
	        kpbm=kpbm&"0"&cstr(day(date))
	      else
	        kpbm=kpbm&cstr(day(date))
	      end if
	      rs.open "select bh from fmc where bh like'" & kpbm & "%' order by bh desc", conn,1,1
	      if rs.recordcount=0 then
	        kpbm=kpbm&"0001"
	      else
	        rs.movefirst
	        st=cstr(cint(right(rs("bh"),4))+1)
	        for i=len(st) to 3
	          st="0"&st
	        next
            kpbm=kpbm&st
	      end if
          rs.close
          unit_name=""
          rs.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"'",conn,1,1
          if rs.recordcount>0 then
            unit_name=unit_name+"["+rs("unit_name")+"]"
            rs.close
            rs.open "select * from unit where left(unit_code," & unit_len2 & ")='" & left(unit_code,unit_len2) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00'",conn,1,1
            if rs.recordcount>0 then
              unit_name=unit_name+"["+rs("unit_name")+"]"
              rs.close
              rs.open "select * from unit where unit_code='"+unit_code+"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00'",conn,1,1
              if rs.recordcount>0 then
                unit_name=unit_name+"["+rs("unit_name")+"]"
              end if
            end if
          end if
          rs.close
	      rs.open "fmc",conn,1,3
          rs.addnew
          rs("bh")=kpbm
          rs("xm")=request("xm")
          rs("xdw")=unit_name
          rs("unit_code")=unit_code
          rs("zw")=request("zw")
          rs("gz")=request("gz")
          rs("sfgrgb")=request("sfgrgb")
          rs("mz")=request("mz")
          rs("xb")=request("xb")
          rs("jg")=request("jg")
          rs("csly")=request("csly")
          rs("wfcdbysj")=request("wfcdbysj")
          rs("wfcdyx")=request("wfcdyx")
          rs("wfcdzy")=request("wfcdzy")
          rs("wfcdxl")=request("wfcdxl")
          rs("rwly")=request("rwly")
          rs("dty")=request("dty")
          rs("sfzh")=request("sfzh")
          rs("zfzh")=request("zfzh")
          rs("zp")=getFileName(zp)'zp
          rs("jc")=request("jc")
          rs("dxr")=username
          rs("czshj")=czshj
          rs.update
          rs.close
          '保存增加花名册
          conn.execute("insert into olog (shj,username,czms,bz) values ('"&now()&"','"&username&"','增加花名册："&unit_name&","&request("xm")&"','ZJFMC')")
	      set rs=nothing
	      closedb()
	      'Response.Redirect "addfmc.asp?mode=1"
	    end if
      else
	    opendb()
        set rs=server.createobject("adodb.recordset")
        unit_name=""
        rs.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"'",conn,1,1
        if rs.recordcount>0 then
          unit_name=unit_name+"["+rs("unit_name")+"]"
          rs.close
          rs.open "select * from unit where left(unit_code," & unit_len2 & ")='" & left(unit_code,unit_len2) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00'",conn,1,1
          if rs.recordcount>0 then
            unit_name=unit_name+"["+rs("unit_name")+"]"
            rs.close
            rs.open "select * from unit where unit_code='"+unit_code+"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00'",conn,1,1
            if rs.recordcount>0 then
              unit_name=unit_name+"["+rs("unit_name")+"]"
            end if
          end if
        end if
        rs.close
	    conn.execute("update fmc set xm='"&request("xm")&"',xdw='"&unit_name&"',unit_code='"&unit_code&"',zw='"&request("zw")&"',gz='"&request("gz")&"',sfgrgb='"&request("sfgrgb")&"',mz='"&request("mz")&"',xb='"&request("xb")&"',jg='"&request("jg")&"',csly='"&request("csly")&"',wfcdbysj='"&request("wfcdbysj")&"',wfcdyx='"&request("wfcdyx")&"',wfcdzy='"&request("wfcdzy")&"',wfcdxl='"&request("wfcdxl")&"',rwly='"&request("rwly")&"',dty='"&request("dty")&"',dxr='"&username&"',sfzh='"&request("sfzh")&"',zfzh='"&request("zfzh")&"',jc='"&request("jc")&"',zp='"&getFileName(zp)&"',czshj='"&czshj&"' where bh='"&request("odq") &"'")
        '保存修改案件日志
        conn.execute("insert into olog (shj,username,czms,bz) values ('"&now()&"','"&username&"','修改花名册："&unit_name&","&request("xm")&"','XGFMC')")
	    closedb()
	    'Response.Redirect "addfmc.asp?mode=1"
      end if
    end if
  else
    ShowInputForm1 mode,""
  end if

elseif mode=4 then
  '删除确认
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="addfmc.asp?mode=1">返回</a>]
     </td>
    </tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除这个花名册“<%=request("xm")%>”？
      <br><br>
      [<a href="addfmc.asp?mode=7&dq=<%=request("dq")%>&xm=<%=request("xm")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="addfmc.asp?mode=1">算了</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()
elseif mode=5 then
  '搜索 

elseif mode=7 then
  'delete
  opendb()
  conn.execute "delete from fmc where bh='" + request("dq")+"'"
  '保存删除案件日志
  conn.execute("insert into olog (shj,username,czms,bz) values ('"&now()&"','"&username&"','删除花名册："&request("xm")&"','SCFMC')")
  closedb()
  delaySecond(2)
  Response.Redirect ("addfmc.asp?mode=1&unit_code=" & request("unit_code"))
elseif mode=102 then
  ShowInputForm1 2,""
elseif mode=103 then
  ShowInputForm1 3,""
end if
%>    