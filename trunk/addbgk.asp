
<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if session("username")=""  or instr(session("power"),",1,")=0 then
  Response.Redirect("notlogin.asp")
end if
%>

<!--#include file="fcommon.asp"-->
<!--#include file="dtp.asp"-->
<%
dim conn, mode, username, rs, sql,rs1,rs2,rs3,rs4,rsMX, errmsg, founderror, s, t, i, fl, dq,odq, cpage,kpbm,st
dim unit_code,firstitem_code,seconditem_code,mode_code,FRMgzgk,FRMitem,FRMmode,FRMkssj,FRMjssj,FRMcx,czshj,explain,i_item
dim shj1,shj2,tjbbname,zhs

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
if not isempty(request("tjbbname")) then
    tjbbname = request("tjbbname")
else
    tjbbname = ""
end if

if not isempty(request("firstitem_code")) then
    firstitem_code = request("firstitem_code")
else
    firstitem_code = ""
end if

if not isempty(request("seconditem_code")) then
    seconditem_code = request("seconditem_code")
else
    seconditem_code = ""
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
  <title>报告卡</title>
  <link rel="stylesheet" type="text/css" href="./main.css">
  </head>
<script language="javascript">
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

function check_form() 
{var word
 //alert(document.form1.odq11.value);
 if(trim(document.form1.shj1.value)=="")
   {alert("请输入日期。");
    return false;
   }
 for(var i=0;i<document.form1.odq11.value;i++)
 {
  //alert(document.form1.ajlbv[i].value );//hjhedit20040420将多个文本框对象的名称都设为一样的，在Java函数访问具体一个文本框对象时，document.窗体名.对象名[下标].value
  if(trim(document.form1.ajlbv[i].value)=="")
    {alert("请输入"+document.form1.ajlb_name[i].value+"。");
     return false;
    } 
  }
  return true;
}

</script>  
  <body leftmargin="0" topmargin="0">
<%noRightClick()
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
    set rs=server.createobject("adodb.recordset")
    set rs1=server.createobject("adodb.recordset")
    set rs2=server.createobject("adodb.recordset")
    set rs3=server.createobject("adodb.recordset")
    set rs4=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(request("tjbb"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
    %>
    <form name="form1" method="post" onsubmit="return check_form()" action="addbgk.asp?mode=2&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    set rs1=server.createobject("adodb.recordset")
    set rs2=server.createobject("adodb.recordset")
    set rs3=server.createobject("adodb.recordset")
    set rs4=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs.open "select * from edzdjb" +left(request("tjbb"),ajlb_len1) +" where bh='" + request("odq") + "'", conn, 1, 1
    rs1.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(request("tjbb"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
    %>
    <form name="form1" method="post" onsubmit="return check_form()" action="addbgk.asp?mode=3&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>&odq=<%=request("odq")%>">
  <%end if%>
  <table width="530" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b><%=tjbbname%>报告卡</b></td>
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
          请填写报告卡信息内容。
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
            <td rowspan="1" colspan="4" bgcolor="#FFFFFF" align="center">日期</td>
            <td align="left" bgcolor="#FFFFFF" width=120>
              <%if mode=2 then%>
                <input type="text" name="shj1" size="10" maxlength="12" readonly value='<%=todatestr(shj1)%>'>
              <%else%>
                <input type="text" name="shj1" size="10" maxlength="12" readonly value='<%=todatestr(rs("shj1"))%>'>
              <%end if%>
              <A onclick="show_cele_date(change1,'','',shj1)"><IMG align=top border=0 height=25 name=change1 src="images\calendar.gif" width=26></A>
            </td>
          </tr>
          <%i_item=0
          i=0
          while not rs1.eof 
            fl=false
            rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(rs1("ajlb_code"),ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
            if rsMX.recordcount=0 then
              if not fl then
                response.write "<tr>"
                fl=true
              end if
              i_item=i_item+1%>
              <td rowspan="1" colspan="4" bgcolor="#FFFFFF" align="center"><%=rs1("ajlb_name")%></td>
              <td bgcolor="#FFFFFF">
                <input name="ajlb_name" type="hidden" value="<%=rs1("ajlb_name")%>"><!--用于获取当前对应文本框所要输入的项目名称-->
                <input name="ajlb_code" type="hidden" value="<%=rs1("ajlb_code")%>"><!--用于获取当前对应文本框所要输入的项目代码-->
                <%if mode=2 then%>
                  <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                <%else
                  rs2.open "select * from edzdjb_x" +left(request("tjbb"),ajlb_len1) +" where bh='"+rs("bh")+"' and ajlb_code='" & rs1("ajlb_code") &"'", conn, 1, 1
                  if rs2.eof then%>
	            <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                  <%else%>
                    <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='<%=rs2("ajlbV")%>' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                  <%end if
                  rs2.close
                end if%>
              </td>
              </tr>
            <%fl=false
            else
              zhs=0
              rsMX.movefirst
              do while not rsMX.eof
                rs3.open "select * from ajlb where left(ajlb_code," & ajlb_len3 & ")='" & left(rsMX("ajlb_code"),ajlb_len3) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' order by ajlb_sxh", conn, 1, 1
                if rs3.recordcount=0 then
                  zhs=zhs+1
                else
                  do while not rs3.eof 
                    rs4.open "select * from ajlb where left(ajlb_code," & ajlb_len4 & ")='" & left(rs3("ajlb_code"),ajlb_len4) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len5) & ")='" & ajlb_str5 &"' and mid(ajlb_code,"& (ajlb_len4+1) & "," & (ajlb_len5-ajlb_len4) & ")<>'00' order by ajlb_sxh", conn, 1, 1
                    if rs4.recordcount=0 then
                      zhs=zhs+1
                    else
                      zhs=zhs+rs4.recordcount
                    end if
                    rs4.close
                    rs3.movenext
                  loop
                end if
                rs3.close
                rsMX.movenext
              loop
              rsMX.movefirst
              if not fl then
                response.write "<tr>"
                fl=true
              end if
              %>
              <td rowspan="<%=zhs%>" colspan="1" bgcolor="#FFFFFF" align="center"><%=rs1("ajlb_name")%></td>
              <%do while not rsMX.eof
                rs3.open "select * from ajlb where left(ajlb_code," & ajlb_len3 & ")='" & left(rsMX("ajlb_code"),ajlb_len3) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' order by ajlb_sxh", conn, 1, 1
                if rs3.recordcount=0 then
                  if not fl then
                    response.write "<tr>"
                    fl=true
                  end if
                  i_item=i_item+1%>
                  <td rowspan="1" colspan="3" bgcolor="#FFFFFF" align="center"><%=rsMX("ajlb_name")%></td>
                  <td bgcolor="#FFFFFF">
                    <input name="ajlb_name" type="hidden" value="<%=rsMX("ajlb_name")%>"><!--用于获取当前对应文本框所要输入的项目名称-->
                    <input name="ajlb_code" type="hidden" value="<%=rsMX("ajlb_code")%>"><!--用于获取当前对应文本框所要输入的项目代码-->
                    <%if mode=2 then%>
                      <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                    <%else
                      rs2.open "select * from edzdjb_x" +left(request("tjbb"),ajlb_len1) +" where bh='"+rs("bh")+"' and ajlb_code='" & rsMX("ajlb_code") &"'", conn, 1, 1
                      if rs2.eof then%>
                        <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                      <%else%>
                        <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='<%=rs2("ajlbV")%>' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                      <%end if
                      rs2.close
                    end if%>
                  </td>
                  </tr>
                <%fl=false
                else
                  zhs=0
                  do while not rs3.eof 
                    rs4.open "select * from ajlb where left(ajlb_code," & ajlb_len4 & ")='" & left(rs3("ajlb_code"),ajlb_len4) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len5) & ")='" & ajlb_str5 &"' and mid(ajlb_code,"& (ajlb_len4+1) & "," & (ajlb_len5-ajlb_len4) & ")<>'00' order by ajlb_sxh", conn, 1, 1
                    if rs4.recordcount=0 then
                      zhs=zhs+1
                    else
                      zhs=zhs+rs4.recordcount
                    end if
                    rs4.close
                    rs3.movenext
                  loop
                  rs3.movefirst
                  if not fl then
                    response.write "<tr>"
                    fl=true
                  end if
                  %>
                  <td rowspan="<%=zhs%>" colspan="1" bgcolor="#FFFFFF" align="center"><%=rsMX("ajlb_name")%></td>
                  <%do while not rs3.eof
                    rs4.open "select * from ajlb where left(ajlb_code," & ajlb_len4 & ")='" & left(rs3("ajlb_code"),ajlb_len4) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len5) & ")='" & ajlb_str5 &"' and mid(ajlb_code,"& (ajlb_len4+1) & "," & (ajlb_len5-ajlb_len4) & ")<>'00' order by ajlb_sxh", conn, 1, 1
                    if rs4.recordcount=0 then  
                      i_item=i_item+1%>
                      <td rowspan="1" colspan="2" bgcolor="#FFFFFF" align="center"><%=rs3("ajlb_name")%></td>
                      <td bgcolor="#FFFFFF">
                        <input name="ajlb_name" type="hidden" value="<%=rs3("ajlb_name")%>"><!--用于获取当前对应文本框所要输入的项目名称-->
                        <input name="ajlb_code" type="hidden" value="<%=rs3("ajlb_code")%>"><!--用于获取当前对应文本框所要输入的项目代码-->
                        <%if mode=2 then%>
                          <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                        <%else
                          rs2.open "select * from edzdjb_x" +left(request("tjbb"),ajlb_len1) +" where bh='"+rs("bh")+"' and ajlb_code='" & rs3("ajlb_code") &"'", conn, 1, 1
                          if rs2.eof then%>
                            <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                          <%else%>
                            <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='<%=rs2("ajlbV")%>' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                          <%end if
                          rs2.close
                        end if%>
                      </td>
                      </tr>
                    <%fl=false
                    else
                      %>
                      <td rowspan="<%=rs4.recordcount%>" bgcolor="#FFFFFF" align="center"><%=rsMX("ajlb_name")%></td>
                      <%do while not rs4.eof
                        if not fl then
                          response.write "<tr>"
                          fl=true
                        end if
                        i_item=i_item+1%>
                        <td rowspan="1" colspan="1" bgcolor="#FFFFFF" align="center"><%=rs4("ajlb_name")%></td>
                        <td bgcolor="#FFFFFF">
                          <input name="ajlb_name" type="hidden" value="<%=rs4("ajlb_name")%>"><!--用于获取当前对应文本框所要输入的项目名称-->
                          <input name="ajlb_code" type="hidden" value="<%=rs4("ajlb_code")%>"><!--用于获取当前对应文本框所要输入的项目代码-->
                          <%if mode=2 then%>
                            <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                          <%else
                            rs2.open "select * from edzdjb_x" +left(request("tjbb"),ajlb_len1) +" where bh='"+rs("bh")+"' and ajlb_code='" & rs4("ajlb_code") &"'", conn, 1, 1
                            if rs2.eof then%>
                              <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                            <%else%>
                              <input name="ajlbv" type="text" size=7 maxlength=7 class="smallInput" value='<%=rs2("ajlbV")%>' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 && window.event.keyCode!=9 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
                            <%end if
                            rs2.close
                          end if%>
                          </td>
                          </tr>
                        <%fl=false
                        rs4.movenext
                      loop
                    end if
                    rs4.close
                    rs3.movenext
                  loop
                end if
                rs3.close
                rsMX.movenext
              loop
            end if
            rsMX.close
            rs1.movenext
          wend%>
          <tr>
            <td colspan="3" bgcolor="#FFFFFF">
            <input name="odq11" type="hidden" value="<%=i_item%>">
            </td>
          </tr>
        </table>
        <p> 
          <input class="buttonface" type="submit" name="Submit" value=" 提 交 ">
          &nbsp; 
          <INPUT class="buttonface" type=reset onclick="{if(confirm('该项操作要清除全部的内容，您确定要清除吗?')){return true;}return false;}" value=" 重 写 " id=reset1 name=reset1>
        </p>        
      </div></td>
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

  <form method="POST" action="addbgk.asp?mode=5&username=<%=username%>" name="input3">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="addbgk.asp?mode=8&username=<%=username%>">返回</a>]
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
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  set rs1=server.createobject("adodb.recordset")
  rs.open "select * from ajlb where ajlb_code='" + request("tjbb") + "'", conn, 1, 1
  if rs.recordcount>0 then 
    tjbbname=rs("ajlb_name")
  end if
  rs.close
  sql="select * from edzdjb" +left(request("tjbb"),ajlb_len1) +" where unit_code='"+session("unit_code")+"' order by shj1 desc"
  'Response.Write sql
  rs.open sql, conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage>rs.pagecount then cpage=1
    rs.AbsolutePage = cpage
    %>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
      </tr>
      <tr>
        <td valign="bottom">第<%=cstr(cpage)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个[<%=tjbbname%>]记录</td>
        <td align="right">
        [<a href="addbgk.asp?mode=2&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>">添加</a>]
        <%if cpage <> 1 then%>
          [<a href="addbgk.asp?mode=1&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>&page=<%=cstr(cpage-1)%>">上一页</a>]
        <%end if%>
        <%if cpage <> rs.PageCount then%>
          [<a href="addbgk.asp?mode=1&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>&page=<%=cstr(cpage+1)%>">下一页</a>]
        <%end if%>
        <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage then%>
                <option selected value="addbgk.asp?mode=1&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="addbgk.asp?mode=1&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
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
              <td width=10% align=center>时间</td>
              <td width=10% align=center>操作</td>
            </tr>
            <%
            fl = False
            for i = 1 to rs.PageSize
            if not rs.EOF then
              if fl then%>
                <tr bgcolor="#eeeeee">
              <%else%>
                <tr>
              <%end if%>
                <td align=center><%=todatestr(rs("shj1"))%></td>
                <td align=center>
                  <a href="addbgk.asp?mode=3&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>&odq=<%=trim(rs("bh"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="addbgk.asp?mode=4&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>&dq=<%=trim(rs("bh"))%>"><img src="./images/del.gif" border=0></a>
                </td>
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
          [<a href="addbgk.asp?mode=2&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>">添加</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有[<%=tjbbname%>]记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=2 or mode=3 then
  '添加及修改
  if request("shj1") <> "" then
    if not isdate( request.form("shj1")) then
      shj1=""
    else
      shj1=datetostr(request.form("shj1"))
    end if
    shj2=shj1
    czshj=now()
    FoundError=false
    ErrMsg=""
    i_item=1
    if FoundError=true then
      ShowInputForm1 mode,errmsg
    else
      if mode=2 then
	    '判断是否已经存在同一时间段同一单位的工作记录
	    opendb()
	    set rs=server.createobject("adodb.recordset")
	    rs.open "select shj1,shj2 from edzdjb" +left(request("tjbb"),ajlb_len1) +" where unit_code='" & session("unit_code") & "' order by shj1 desc",conn,1,1
        do while not rs.eof
          if not((shj1<rs("shj1") and shj2<rs("shj2")) or (shj1>rs("shj2") and shj2>rs("shj2"))) then
            if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
            ErrMsg = ErrMsg + "时间段重复"
	        FoundError = True
            exit do
          end if
          rs.movenext
        loop
        rs.close
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
	      rs.open "select bh from edzdjb" +left(request("tjbb"),ajlb_len1) +" where bh like'" & kpbm & "%' order by bh desc", conn,1,1
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
	      '先插入表单信息（edzdjb），主要是bh,shj1,shj2,unit_cod,username,czshj
	      conn.execute("insert into edzdjb" +left(request("tjbb"),ajlb_len1) +" (bh,shj1,shj2,unit_code,username,czshj) values ('"&kpbm&"','" & shj1 & "','" & shj2 & "','" & session("unit_code")&"','"&session("username")&"','"&czshj&"')")
	      '后插入各项目的ajlb_code,ajlbV，以bh为连接关健字
	      i_item=1
	      while i_item<=Request.Form("ajlbv").Count 
	        conn.execute("insert into edzdjb_x" +left(request("tjbb"),ajlb_len1) +" (bh,ajlb_code,ajlbV) values ('"&kpbm&"','" & Request.Form("ajlb_code")(i_item) & "','" & Request.Form("ajlbv")(i_item) &"')")
	        i_item=i_item+1
	      wend
	      rs.close
	      set rs=nothing
	      closedb()
	      Response.Redirect "addbgk.asp?mode=1&tjbb="&request("tjbb")&"&tjbbname="&tjbbname&"&kpbm="&kpbm
	    end if
      else
        '判断是否已经存在同一时间段同一单位的工作记录
	    opendb()
	    set rs=server.createobject("adodb.recordset")
	    rs.open "select shj1,shj2 from edzdjb" +left(request("tjbb"),ajlb_len1) +" where unit_code='" & session("unit_code") & "' and bh<>'"+request("odq") & "' order by shj1 desc",conn,1,1
        do while not rs.eof
          if not((shj1<rs("shj1") and shj2<rs("shj2")) or (shj1>rs("shj2") and shj2>rs("shj2"))) then
            if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
            ErrMsg = ErrMsg + "时间段重复"
	        FoundError = True
            exit do
          end if
          rs.movenext
        loop
        rs.close
        if FoundError=true then
	      set rs=nothing
	      closedb()
	      ShowInputForm1 mode,errmsg
	    else
          opendb()
          conn.Execute("delete from edzdjb_x" +left(request("tjbb"),ajlb_len1) +" where bh='"&request("odq") &"'")'先清除旧的数据
          i_item=1
          while i_item<=Request.Form("ajlbv").Count '保存新的责任区各项目数据
	       conn.execute("insert into edzdjb_x" +left(request("tjbb"),ajlb_len1) +" (bh,ajlb_code,ajlbV) values ('"&request("odq") &"','" & Request.Form("ajlb_code")(i_item) & "','" & Request.Form("ajlbv")(i_item) &"')")
           i_item=i_item+1
          wend
          closedb()
          Response.Redirect "addbgk.asp?mode=1&tjbb="&request("tjbb")&"&tjbbname="&tjbbname&"&kpbm="&kpbm
        end if
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
        [<a href="addbgk.asp?mode=1">返回</a>]
     </td>
    </tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除这个[<%=tjbbname%>]报告卡“<%=request("dq")%>”？
      <br><br>
      [<a href="addbgk.asp?mode=7&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>&dq=<%=request("dq")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="addbgk.asp?mode=1&tjbb=<%=request("tjbb")%>&tjbbname=<%=tjbbname%>">算了</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=7 then
  'delete
  opendb()
  conn.execute "delete from edzdjb" +left(request("tjbb"),ajlb_len1) +" where bh='" + request("dq")+"'"
  conn.execute "delete from edzdjb_x" +left(request("tjbb"),ajlb_len1) +" where bh='" + request("dq")+"'"
  closedb()
  delaySecond(2)
  Response.Redirect ("addbgk.asp?mode=1&tjbb="&request("tjbb")&"&tjbbname="&tjbbname&"")
end if
%>    