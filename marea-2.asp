<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if session("username")="" or instr(session("power"),",0,")=0 then
    Response.Redirect "notlogin.asp"
end if
%>

<!--#include file="fcommon.asp"-->

<%
dim conn, mode, username, rs, sql, errmsg, founderror, s, t, i, fl, dq,odq,dq0, dq1,cpage1,cpage2,cpage3,cpage4,cpage5,kpbm,st,dwxh,sfzs,dqcode1,dqcode2,dqcode3,dqcode4,dqcode5,dqname1,dqname2,dqname3,dqname4,dqname5
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
if not isempty(request("dqcode1")) then
    dqcode1 = request("dqcode1")
else
    dqcode1 = ""
end if
if not isempty(request("dqcode2")) then
    dqcode2 = request("dqcode2")
else
    dqcode2 = ""
end if
if not isempty(request("dqcode3")) then
    dqcode3 = request("dqcode3")
else
    dqcode3 = ""
end if
if not isempty(request("dqcode4")) then
    dqcode4 = request("dqcode4")
else
    dqcode4 = ""
end if
if not isempty(request("dqcode5")) then
    dqcode5 = request("dqcode5")
else
    dqcode5 = ""
end if
if not isempty(request("dqname1")) then
    dqname1 = request("dqname1")
else
    dqname1 = ""
end if
if not isempty(request("dqname2")) then
    dqname2 = request("dqname2")
else
    dqname2 = ""
end if
if not isempty(request("dqname3")) then
    dqname3 = request("dqname3")
else
    dqname3 = ""
end if
if not isempty(request("dqname4")) then
    dqname4 = request("dqname4")
else
    dqname4 = ""
end if
if not isempty(request("dqname5")) then
    dqname5 = request("dqname5")
else
    dqname5 = ""
end if
  if not isEmpty(request("page1")) and isnumeric(request("page1")) then
    cpage1 = clng(request("page1"))
  else
    cpage1 = 1
  end if
  if not isEmpty(request("page2")) and isnumeric(request("page2")) then
    cpage2 = clng(request("page2"))
  else
    cpage2 = 1
  end if
  if not isEmpty(request("page3")) and isnumeric(request("page3")) then
    cpage3 = clng(request("page3"))
  else
    cpage3 = 1
  end if
  if not isEmpty(request("page4")) and isnumeric(request("page4")) then
    cpage4 = clng(request("page4"))
  else
    cpage4 = 1
  end if
  if not isEmpty(request("page5")) and isnumeric(request("page5")) then
    cpage5 = clng(request("page5"))
  else
    cpage5 = 1
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
  <title>案件类别管理</title>
  <link rel="stylesheet" type="text/css" href="./main.css">
  </head>

  <script LANUGAGE="JavaScript">
  <!--
  function surfto(list){
   var myindex1=list.selectedIndex;
   if (myindex1!=0 & myindex1!=1){ location.href=list.options[list.selectedIndex].value }
  }
  function goto(list){
   location.href=list.options[list.selectedIndex].value
  }
  //-->
  </script>

  <body>
  <%noRightClick()%>
  <table width="90%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28"><td align="center">
      <%if mode<100 then %>
        <b>报告卡类别设置</b>
      <%elseif mode>100 and mode<200 then %>
        <b>报告卡[<%=request("dqname1")%>]-1类类别设置</b>
      <%elseif mode>200 and mode<300 then %>
        <b>报告卡[<%=request("dqname1")%>]-1类[<%=request("dqname2")%>]-2类类别设置</b>
      <%elseif mode>300 and mode<400 then %>
        <b>报告卡[<%=request("dqname1")%>]-1类[<%=request("dqname2")%>]-2类[<%=request("dqname3")%>]-3类类别设置</b>
      <%elseif mode>400 and mode<500 then %>
        <b>报告卡[<%=request("dqname1")%>]-1类[<%=request("dqname2")%>]-2类[<%=request("dqname3")%>]-3类[<%=request("dqname4")%>]-4类类别设置</b>
      <%end if%>
    </td></tr>
  </table>
  <br>
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
  if mode = 2 then%>
    <form method="POST" action="marea-2.asp?mode=2&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from ajlb where ajlb_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-2.asp?mode=3&page1=<%=cpage1%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-2.asp?mode=1&page1=<%=cpage1%>">返回</a>]
      </td>
    </tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
    <tr><td>
      <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
        <tr>
        <%if Trim(ErrMsg) <> "" then%>
          <td colspan="3"><%=errmsg%></td>
        <%else%>
          <% if mode = 2 then%>
            <td colspan="3">请输入报告卡类别，然后点击“OK”</td>
          <%else%>
            <td colspan="3">请编辑报告卡类别，然后点击“OK”</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>报告卡类别代码&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=dq size=15 maxlength=<%=ajlb_len1%> class="smallInput" value='<%=request("dq")%>'>
            <%else%>
              <input name=dq size=15 maxlength=<%=ajlb_len1%> class="smallInput" value='<%=trim(left(rs("ajlb_code"),ajlb_len1))%>'>
            <%end if%>
            <font color=red>(*)</font>(请输入编号前<%=ajlb_len1%>位,后<%=ajlb_len0-ajlb_len1%>位全为0)
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>报告卡类别名称&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=dq0 size=15 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=15 maxlength=30 class="smallInput" value='<%=trim(rs("ajlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr align="center">
          <td colspan="3"><input class="buttonface" type="submit" value=" 确 定 " id=submit1 name=submit1></td>
        </tr>
      </table>
    </td></tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
  </table>
  </form>
<%
  if mode = 3 then
    rs.close
    set rs=nothing
    closedb()
  end if
  showctail
end sub

sub ShowInputForm101(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 102 then%>
    <form method="POST" action="marea-2.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from ajlb where ajlb_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-2.asp?mode=103&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>">返回</a>]
      </td>
    </tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
    <tr><td>
      <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
        <tr>
        <%if Trim(ErrMsg) <> "" then%>
          <td colspan="3"><%=errmsg%></td>
        <%else%>
          <% if mode = 102 then%>
            <td colspan="3">请输入1类，然后点击“OK”</td>
          <%else%>
            <td colspan="3">请编辑1类，然后点击“OK”</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>1类代码&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=dq size=15 maxlength=<%=ajlb_len3-ajlb_len2%> class="smallInput" value='<%=request("dq")%>'>(前<%=ajlb_len1%>位为<%=left(request("dqcode1"),ajlb_len1)%>,输入后<%=ajlb_len3-ajlb_len2%>位)
            <%else%>
              <input name=dq size=15 maxlength=<%=ajlb_len3-ajlb_len2%> class="smallInput" value='<%=trim(mid(rs("ajlb_code"),ajlb_len1+1,ajlb_len2-ajlb_len1))%>'>(前<%=ajlb_len1%>位为<%=left(request("dqcode1"),ajlb_len1)%>,输入后<%=ajlb_len3-ajlb_len2%>位)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>1类名称&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("ajlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr align="center">
          <td colspan="3"><input class="buttonface" type="submit" value=" 确 定 " id=submit1 name=submit1></td>
        </tr>
      </table>
    </td></tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
  </table>
  </form>
  <%
  if mode = 103 then
    rs.close
    set rs=nothing
    closedb()
  end if
  showctail
end sub

sub ShowInputForm201(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 202 then%>
    <form method="POST" action="marea-2.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from ajlb where ajlb_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-2.asp?mode=203&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>">返回</a>]
      </td>
    </tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
    <tr><td>
      <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
        <tr>
        <%if Trim(ErrMsg) <> "" then%>
          <td colspan="3"><%=errmsg%></td>
        <%else%>
          <% if mode = 202 then%>
            <td colspan="3">请输入2类，然后点击“OK”</td>
          <%else%>
            <td colspan="3">请编辑2类，然后点击“OK”</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>2类代码&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=dq size=15 maxlength=<%=ajlb_len2-ajlb_len1%> class="smallInput" value='<%=request("dq")%>'>(前<%=ajlb_len2%>位为<%=left(request("dqcode2"),ajlb_len2)%>,输入后<%=ajlb_len3%>位)
            <%else%>
              <input name=dq size=15 maxlength=<%=ajlb_len2-ajlb_len1%> class="smallInput" value='<%=trim(mid(rs("ajlb_code"),ajlb_len2+1,ajlb_len3-ajlb_len2))%>'>(前<%=ajlb_len2%>位为<%=left(request("dqcode2"),ajlb_len2)%>,输入后<%=ajlb_len3%>位)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>2类名称&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("ajlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr align="center">
          <td colspan="3"><input class="buttonface" type="submit" value=" 确 定 " id=submit1 name=submit1></td>
        </tr>
      </table>
    </td></tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
  </table>
  </form>
<%
  if mode = 203 then
    rs.close
    set rs=nothing
    closedb()
  end if
  showctail
end sub

sub ShowInputForm301(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 302 then%>
    <form method="POST" action="marea-2.asp?mode=302&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from ajlb where ajlb_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-2.asp?mode=303&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>">返回</a>]
      </td>
    </tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
    <tr><td>
      <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
        <tr>
        <%if Trim(ErrMsg) <> "" then%>
          <td colspan="3"><%=errmsg%></td>
        <%else%>
          <% if mode = 302 then%>
            <td colspan="3">请输入3类，然后点击“OK”</td>
          <%else%>
            <td colspan="3">请编辑3类，然后点击“OK”</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>3类代码&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=302 then%>
              <input name=dq size=15 maxlength=<%=ajlb_len4-ajlb_len3%> class="smallInput" value='<%=request("dq")%>'>(前<%=ajlb_len3%>位为<%=left(request("dqcode3"),ajlb_len3)%>,输入后<%=ajlb_len4-ajlb_len3%>位)
            <%else%>
              <input name=dq size=15 maxlength=<%=ajlb_len4-ajlb_len3%> class="smallInput" value='<%=trim(mid(rs("ajlb_code"),ajlb_len3+1,ajlb_len4-ajlb_len3))%>'>(前<%=ajlb_len3%>位为<%=left(request("dqcode3"),ajlb_len3)%>,输入后<%=ajlb_len4-ajlb_len3%>位)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>3类名称&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=302 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("ajlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr align="center">
          <td colspan="3"><input class="buttonface" type="submit" value=" 确 定 " id=submit1 name=submit1></td>
        </tr>
      </table>
    </td></tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
  </table>
  </form>
<%
  if mode = 303 then
    rs.close
    set rs=nothing
    closedb()
  end if
  showctail
end sub

sub ShowInputForm401(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 402 then%>
    <form method="POST" action="marea-2.asp?mode=402&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from ajlb where ajlb_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-2.asp?mode=403&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cpage5%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-2.asp?mode=401&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cpage5%>">返回</a>]
      </td>
    </tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
    <tr><td>
      <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
        <tr>
        <%if Trim(ErrMsg) <> "" then%>
          <td colspan="3"><%=errmsg%></td>
        <%else%>
          <% if mode = 402 then%>
            <td colspan="3">请输入4类，然后点击“OK”</td>
          <%else%>
            <td colspan="3">请编辑4类，然后点击“OK”</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>4类代码&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=402 then%>
              <input name=dq size=15 maxlength=<%=ajlb_len4-ajlb_len3%> class="smallInput" value='<%=request("dq")%>'>(前<%=ajlb_len3%>位为<%=left(request("dqcode3"),ajlb_len3)%>,输入后<%=ajlb_len4-ajlb_len3%>位)
            <%else%>
              <input name=dq size=15 maxlength=<%=ajlb_len4-ajlb_len3%> class="smallInput" value='<%=trim(mid(rs("ajlb_code"),ajlb_len3+1,ajlb_len4-ajlb_len3))%>'>(前<%=ajlb_len3%>位为<%=left(request("dqcode3"),ajlb_len3)%>,输入后<%=ajlb_len4-ajlb_len3%>位)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>4类名称&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=402 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("ajlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr align="center">
          <td colspan="3"><input class="buttonface" type="submit" value=" 确 定 " id=submit1 name=submit1></td>
        </tr>
      </table>
    </td></tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
  </table>
  </form>
<%
  if mode = 403 then
    rs.close
    set rs=nothing
    closedb()
  end if
  showctail
end sub

if mode=1 then
  '报告卡类别显示 
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  'response.write "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh"
  rs.open "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage1>rs.pagecount then cpage1=1
    rs.AbsolutePage = cpage1
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">第<%=cstr(cpage1)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个大类项目</td>
          <td align="right">
          [<a href="marea-2.asp?mode=2">添加</a>]
          <%if cpage1 <> 1 then%>
            [<a href="marea-2.asp?mode=1&page1=<%=cstr(cpage1-1)%>">上一页</a>]
          <%end if%>
          <%if cpage1 <> rs.PageCount then%>
            [<a href="marea-2.asp?mode=1&page1=<%=cstr(cpage1+1)%>">下一页</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage1 then%>
                <option selected value="marea-2.asp?mode=1&page1=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="marea-2.asp?mode=1&page1=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>报告卡代码</td>
              <td width=40% align=center>报告卡类别名称</td>
              <td width=50% align=center>操作</td>
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
                <td align=center><%=trim(rs("ajlb_code"))%></td>
                <td align=center><%=trim(rs("ajlb_name"))%><a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=trim(rs("ajlb_code"))%>&dqname1=<%=trim(rs("ajlb_name"))%>">（<font color="#FF0000">1类</font>）</a></td>
                <td align=center>
                  <a href="marea-2.asp?mode=3&page1=<%=cpage1%>&odq=<%=trim(rs("ajlb_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-2.asp?mode=4&page1=<%=cpage1%>&dq=<%=trim(rs("ajlb_code"))%>&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("ajlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=8&page1=<%=cpage1%>&dq=<%=trim(rs("ajlb_code"))%>&sort=up&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("ajlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=8&page1=<%=cpage1%>&dq=<%=trim(rs("ajlb_code"))%>&sort=down&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
                  <%end if%>
                </td>
              </tr>
              <%rs.MoveNext
              fl = not fl
              end if
            next%>
          </table>
        </td></tr>
      </table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          [<a href="marea-2.asp?mode=2">添加</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有报告卡类别记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()
elseif mode=101 then
  '1类显示
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  'Response.Write("select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh")
  rs.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage2>rs.pagecount then cpage2=1
    rs.AbsolutePage = cpage2
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">第<%=cstr(cpage2)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个1类项目</td>
          <td align="right">
          [<a href="marea-2.asp?mode=1&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">报告卡列表</a>]&nbsp;
          [<a href="marea-2.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">添加</a>]
          <%if cpage2 <> 1 then%>
            [<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(cpage2-1)%>">上一页</a>]
          <%end if%>
          <%if cpage2 <> rs.PageCount then%>
            [<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(cpage2+1)%>">下一页</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2" onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage2 then%>
                <option selected value="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>1类代码</td>
              <td width=40% align=center>1类名称</td>
              <td width=50% align=center>操作</td>
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
                <td align=center><%=trim(rs("ajlb_code"))%></td>
                <td align=center><%=trim(rs("ajlb_name"))%><a href="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=trim(request("dqcode1"))%>&dqname1=<%=trim(request("dqname1"))%>&page2=<%=cpage2%>&dqcode2=<%=trim(rs("ajlb_code"))%>&dqname2=<%=trim(rs("ajlb_name"))%>">（<font color="#FF0000">2类</font>）</a></td>
                <td align=center>
                  <a href="marea-2.asp?mode=103&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&odq=<%=trim(rs("ajlb_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-2.asp?mode=104&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("ajlb_code"))%>&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("ajlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=108&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("ajlb_code"))%>&sort=up&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("ajlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=108&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("ajlb_code"))%>&sort=down&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
                  <%end if%>
                </td>
              </tr>
              <%rs.MoveNext
              fl = not fl
              end if
            next%>
          </table>
        </td></tr>
      </table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          [<a href="marea-2.asp?mode=1&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">报告卡列表</a>]&nbsp;
          [<a href="marea-2.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">添加</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有1类记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=201 then
  '2类显示
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(request("dqcode2"),ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage3>rs.pagecount then cpage3=1
    rs.AbsolutePage = cpage3
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">第<%=cstr(cpage2)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个2类项目</td>
          <td align="right">
          [<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>">1类列表</a>]&nbsp;
          [<a href="marea-2.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>">添加</a>]
          <%if cpage3 <> 1 then%>
            [<a href="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cstr(cpage3-1)%>">上一页</a>]
          <%end if%>
          <%if cpage3 <> rs.PageCount then%>
            [<a href="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cstr(cpage3+1)%>">下一页</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage3 then%>
                <option selected value="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>2类代码</td>
              <td width=40% align=center>2类名称</td>
              <td width=50% align=center>操作</td>
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
                <td align=center><%=trim(rs("ajlb_code"))%></td>
                <td align=center><%=trim(rs("ajlb_name"))%><a href="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=trim(request("dqcode1"))%>&dqname1=<%=trim(request("dqname1"))%>&page2=<%=cpage2%>&dqcode2=<%=trim(request("dqcode2"))%>&dqname2=<%=trim(request("dqname2"))%>&page3=<%=cpage3%>&dqcode3=<%=trim(rs("ajlb_code"))%>&dqname3=<%=trim(rs("ajlb_name"))%>">（<font color="#FF0000">3类</font>）</a></td>
                <td align=center>
                  <a href="marea-2.asp?mode=203&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&odq=<%=trim(rs("ajlb_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-2.asp?mode=204&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("ajlb_code"))%>&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("ajlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=208&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("ajlb_code"))%>&sort=up&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("ajlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=208&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("ajlb_code"))%>&sort=down&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
                  <%end if%>
                </td>
              </tr>
              <%rs.MoveNext
              fl = not fl
              end if
            next%>
          </table>
        </td></tr>
      </table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          [<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>">1类列表</a>]&nbsp;
          [<a href="marea-2.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>">添加</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有2类记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=301 then
  '3类显示
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  rs.open "select * from ajlb where left(ajlb_code," & ajlb_len3 & ")='" & left(request("dqcode3"),ajlb_len3) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' order by ajlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage4>rs.pagecount then cpage4=1
    rs.AbsolutePage = cpage4
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">第<%=cstr(cpage2)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个3类项目</td>
          <td align="right">
          [<a href="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>">2类列表</a>]&nbsp;
          [<a href="marea-2.asp?mode=302&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>">添加</a>]
          <%if cpage4 <> 1 then%>
            [<a href="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cstr(cpage4-1)%>">上一页</a>]
          <%end if%>
          <%if cpage4 <> rs.PageCount then%>
            [<a href="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cstr(cpage4+1)%>">下一页</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage4 then%>
                <option selected value="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>3类代码</td>
              <td width=40% align=center>3类名称</td>
              <td width=50% align=center>操作</td>
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
                <td align=center><%=trim(rs("ajlb_code"))%></td>
                <td align=center><%=trim(rs("ajlb_name"))%><a href="marea-2.asp?mode=401&page1=<%=cpage1%>&dqcode1=<%=trim(request("dqcode1"))%>&dqname1=<%=trim(request("dqname1"))%>&page2=<%=cpage2%>&dqcode2=<%=trim(request("dqcode2"))%>&dqname2=<%=trim(request("dqname2"))%>&page3=<%=cpage3%>&dqcode3=<%=trim(request("dqcode3"))%>&dqname3=<%=trim(request("dqname3"))%>&page4=<%=cpage4%>&dqcode4=<%=trim(rs("ajlb_code"))%>&dqname4=<%=trim(rs("ajlb_name"))%>">（<font color="#FF0000">4类</font>）</a></td>
                <td align=center>
                  <a href="marea-2.asp?mode=303&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&odq=<%=trim(rs("ajlb_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-2.asp?mode=304&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dq=<%=trim(rs("ajlb_code"))%>&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("ajlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=308&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dq=<%=trim(rs("ajlb_code"))%>&sort=up&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("ajlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=308&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dq=<%=trim(rs("ajlb_code"))%>&sort=down&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
                  <%end if%>
                </td>
              </tr>
              <%rs.MoveNext
              fl = not fl
              end if
            next%>
          </table>
        </td></tr>
      </table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          [<a href="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>">2类列表</a>]&nbsp;
          [<a href="marea-2.asp?mode=302&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>">添加</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有3类记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=401 then
  '4类显示
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  rs.open "select * from ajlb where left(ajlb_code," & ajlb_len4 & ")='" & left(request("dqcode4"),ajlb_len4) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len5) & ")='" & ajlb_str5 &"' and mid(ajlb_code,"& (ajlb_len4+1) & "," & (ajlb_len5-ajlb_len4) & ")<>'00' order by ajlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage4>rs.pagecount then cpage4=1
    rs.AbsolutePage = cpage4
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">第<%=cstr(cpage2)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个4类项目</td>
          <td align="right">
          [<a href="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>">3类列表</a>]&nbsp;
          [<a href="marea-2.asp?mode=402&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>">添加</a>]
          <%if cpage4 <> 1 then%>
            [<a href="marea-2.asp?mode=401&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cstr(cpage5-1)%>">上一页</a>]
          <%end if%>
          <%if cpage4 <> rs.PageCount then%>
            [<a href="marea-2.asp?mode=401&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cstr(cpage5+1)%>">下一页</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage4 then%>
                <option selected value="marea-2.asp?mode=401&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="marea-2.asp?mode=401&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>4类代码</td>
              <td width=40% align=center>4类名称</td>
              <td width=50% align=center>操作</td>
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
                <td align=center><%=trim(rs("ajlb_code"))%></td>
                <td align=center><%=trim(rs("ajlb_name"))%></td>
                <td align=center>
                  <a href="marea-2.asp?mode=403&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cpage5%>&odq=<%=trim(rs("ajlb_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-2.asp?mode=404&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cpage5%>&dq=<%=trim(rs("ajlb_code"))%>&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("ajlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=408&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cpage5%>&dq=<%=trim(rs("ajlb_code"))%>&sort=up&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("ajlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-2.asp?mode=408&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cpage5%>&dq=<%=trim(rs("ajlb_code"))%>&sort=down&dwxh=<%=trim(rs("ajlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
                  <%end if%>
                </td>
              </tr>
              <%rs.MoveNext
              fl = not fl
              end if
            next%>
          </table>
        </td></tr>
      </table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          [<a href="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>">3类列表</a>]&nbsp;
          [<a href="marea-2.asp?mode=402&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>">添加</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有4类记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=2 or mode=3 then
  '报告卡类别添加及修改
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to ajlb_len1-1
      dq="0"+cstr(dq)
    next 
    for i=len(dq) to ajlb_len0-1
      dq=cstr(dq)+"0"
    next 
    dq0 = trim(request("dq0"))
    'response.write dq
    if mode=2 then
      if dq = "" then
        ErrMsg="请输入报告卡类别代码"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "报告卡类别代码重复"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="请输入报告卡类别名称"
        foundError=True
      end if
    else
      '看改过的用户名是否存在
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "报告卡类别代码重复"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm1 mode,errmsg
    else
      if mode = 2 then
        '是添加
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "ajlb", conn, 1, 3
        rs.addnew
        rs("ajlb_code")=dq
        rs("ajlb_name")=dq0
        rs("ajlb_sxh")=dwxh
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-2.asp?mode=1"
      else
        opendb()
        conn.Execute "update ajlb set ajlb_code='"+dq+"',ajlb_name='"+dq0+"' where ajlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-2.asp?mode=1&page1=" & cpage1
      end if
    end if
  else
      ShowInputForm1 mode,""
  end if

elseif mode=102 or mode=103 then
  '1类添加及修改
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to ajlb_len2-ajlb_len1-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode1"),ajlb_len1)+ dq
    for i=len(dq) to ajlb_len0-1
      dq=cstr(dq)+"0"
    next
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    if mode=102 then
      if dq = "" then
        ErrMsg="请输入1类代码"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "1类代码重复"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select ajlb_name from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="请输入1类名称"
        foundError=True
      end if
    else
      '看改过的用户名是否存在
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "1类代码重复"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm101 mode,errmsg
    else
      if mode = 102 then
        '是添加
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "ajlb", conn, 1, 3
        rs.addnew
        rs("ajlb_code")=dq
        rs("ajlb_name")=dq0
        rs("ajlb_sxh")=dwxh
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-2.asp?mode=101&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")
      else
        opendb()
        conn.Execute "update ajlb set ajlb_code='"+dq+"',ajlb_name='"+dq0+"' where ajlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-2.asp?mode=101&page1="& cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2
      end if
    end if
  else
      ShowInputForm101 mode,""
  end if

elseif mode=202 or mode=203 then
  '2类添加及修改
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to ajlb_len3-ajlb_len2-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode2"),ajlb_len2)+ dq
    for i=len(dq) to ajlb_len0-1
      dq=cstr(dq)+"0"
    next
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    if mode=202 then
      if dq = "" then
        ErrMsg="请输入2类代码"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "2类代码重复"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select ajlb_name from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(request("dqcode2"),ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="请输入2类名称"
        foundError=True
      end if
    else
      '看改过的用户名是否存在
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "2类代码重复"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm201 mode,errmsg
    else
      if mode = 202 then
        '是添加
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "ajlb", conn, 1, 3
        rs.addnew
        rs("ajlb_code")=dq
        rs("ajlb_name")=dq0
        rs("ajlb_sxh")=dwxh
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-2.asp?mode=201&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")& "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2")
      else
        opendb()
        conn.Execute "update ajlb set ajlb_code='"+dq+"',ajlb_name='"+dq0+"' where ajlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-2.asp?mode=201&page1="& cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3
      end if
    end if
  else
      ShowInputForm201 mode,""
  end if

elseif mode=302 or mode=303 then
  '3类添加及修改
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to ajlb_len4-ajlb_len3-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode3"),ajlb_len3)+ dq
    for i=len(dq) to ajlb_len0-1
      dq=cstr(dq)+"0"
    next
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    if mode=302 then
      if dq = "" then
        ErrMsg="请输入3类代码"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "3类代码重复"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select ajlb_name from ajlb where left(ajlb_code," & ajlb_len3 & ")='" & left(request("dqcode3"),ajlb_len3) &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' order by ajlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="请输入3类名称"
        foundError=True
      end if
    else
      '看改过的用户名是否存在
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "3类代码重复"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm301 mode,errmsg
    else
      if mode = 302 then
        '是添加
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "ajlb", conn, 1, 3
        rs.addnew
        rs("ajlb_code")=dq
        rs("ajlb_name")=dq0
        rs("ajlb_sxh")=dwxh
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-2.asp?mode=301&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")& "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2")& "&page3=" & cpage3& "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3")
      else
        opendb()
        conn.Execute "update ajlb set ajlb_code='"+dq+"',ajlb_name='"+dq0+"' where ajlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-2.asp?mode=301&page1="& cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3& "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3") & "&page4=" & cpage4
      end if
    end if
  else
      ShowInputForm301 mode,""
  end if

elseif mode=402 or mode=403 then
  '4类添加及修改
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to ajlb_len5-ajlb_len4-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode4"),ajlb_len4)+ dq
    for i=len(dq) to ajlb_len0-1
      dq=cstr(dq)+"0"
    next
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    if mode=402 then
      if dq = "" then
        ErrMsg="请输入4类代码"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "4类代码重复"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select ajlb_name from ajlb where left(ajlb_code," & ajlb_len4 & ")='" & left(request("dqcode4"),ajlb_len4) &"' and mid(ajlb_code,"& (ajlb_len4+1) & "," & (ajlb_len5-ajlb_len4) & ")<>'00' order by ajlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="请输入4类名称"
        foundError=True
      end if
    else
      '看改过的用户名是否存在
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "4类代码重复"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm301 mode,errmsg
    else
      if mode = 402 then
        '是添加
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "ajlb", conn, 1, 3
        rs.addnew
        rs("ajlb_code")=dq
        rs("ajlb_name")=dq0
        rs("ajlb_sxh")=dwxh
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-2.asp?mode=401&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")& "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2")& "&page3=" & cpage3& "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3")& "&page4=" & cpage4& "&dqcode4="+request("dqcode4") & "&dqname4=" & request("dqname4")
      else
        opendb()
        conn.Execute "update ajlb set ajlb_code='"+dq+"',ajlb_name='"+dq0+"' where ajlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-2.asp?mode=401&page1="& cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3& "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3") & "&page4=" & cpage4& "&dqcode4="+request("dqcode4") & "&dqname4=" & request("dqname4")& "&page5=" & cpage5
      end if
    end if
  else
      ShowInputForm401 mode,""
  end if

elseif mode=4 then
  '报告卡类别删除确认
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除报告卡类别“<%=request("dq")%>”？
      <br><br>
      [<a href="marea-2.asp?mode=7&page1=<%=cpage1%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-2.asp?mode=1&page1=<%=cpage1%>">算了</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=104 then
  '1类删除确认
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除1类“<%=request("dq")%>”？
      <br><br>
      [<a href="marea-2.asp?mode=107&page1=<%=cpage1%>&page2=<%=cpage2%>&dq=<%=request("dq")%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&dwxh=<%=request("dwxh")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&page2=<%=cpage2%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">算了</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=204 then
  '2类删除确认
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除2类“<%=request("dq")%>”？
      <br><br>
      [<a href="marea-2.asp?mode=207&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-2.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>">算了</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=304 then
  '3类删除确认
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除3类“<%=request("dq")%>”？
      <br><br>
      [<a href="marea-2.asp?mode=307&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-2.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>">算了</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=404 then
  '4类删除确认
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除4类“<%=request("dq")%>”？
      <br><br>
      [<a href="marea-2.asp?mode=407&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cpage5%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-2.asp?mode=401&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&page4=<%=cpage4%>&dqcode4=<%=request("dqcode4")%>&dqname4=<%=request("dqname4")%>&page5=<%=cpage5%>">算了</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=7 then
  '报告卡类别delete
  opendb()
  conn.execute "delete from ajlb where ajlb_code like '" + left(request("dq"),ajlb_len1)+"%'"'清除本身大类及所属的中类和小类
  conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' and ajlb_sxh>=" & request("dwxh")' 后面的顺序号往前推
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=1&page1=" & cpage1)

elseif mode=107 then
  '1类delete
  opendb()
  conn.execute "delete from ajlb where ajlb_code like'" + left(request("dq"),ajlb_len2)+"%'"'清除本身中类及所属的小类
  conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' and ajlb_sxh>=" & request("dwxh")' 后面的顺序号往前推
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=101&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2)  

elseif mode=207 then
  '2类delete
  opendb()
  conn.execute "delete from ajlb where ajlb_code like'" + left(request("dq"),ajlb_len3)+"%'"'清除本身小类
  conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where left(ajlb_code," & ajlb_len2 & ")='" & left(request("dqcode2"),ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' and ajlb_sxh>=" & request("dwxh")' 后面的顺序号往前推
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=201&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") &"&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3)

elseif mode=307 then
  '3类delete
  opendb()
  conn.execute "delete from ajlb where ajlb_code like'" + left(request("dq"),ajlb_len4)+"%'"'清除本身案件类别
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=301&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") &"&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") &"&page3=" & cpage3 & "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3") & "&page4=" & cpage4)

elseif mode=407 then
  '4类delete
  opendb()
  conn.execute "delete from ajlb where ajlb_code like'" + left(request("dq"),ajlb_len5)+"%'"'清除本身案件类别
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=401&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") &"&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") &"&page3=" & cpage3 & "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3") &"&page4=" & cpage4 & "&dqcode4="+request("dqcode4") & "&dqname4=" & request("dqname4") & "&page5=" & cpage5)

elseif mode=8 then
  'delete 报告卡类别上移/下移
  opendb()
  if request("sort")="up" then'上移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' and ajlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where ajlb_code='" + request("dq")+"'"
  else'下移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' and ajlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where ajlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=1&page1=" & cpage1)

elseif mode=108 then
  'delete 1类上移/下移
  opendb()
  if request("sort")="up" then'上移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where ajlb_code='" + request("dq")+"'"
  else'下移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where ajlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=101&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2)

elseif mode=208 then
  'delete 2类上移/下移
  opendb()
  if request("sort")="up" then'上移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where left(ajlb_code," & ajlb_len2 & ")='" & left(request("dqcode2"),ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where ajlb_code='" + request("dq")+"'"
  else'下移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where left(ajlb_code," & ajlb_len2 & ")='" & left(request("dqcode2"),ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where ajlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=201&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3)

elseif mode=308 then
  'delete 3类上移/下移
  opendb()
  if request("sort")="up" then'上移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where left(ajlb_code," & ajlb_len3 & ")='" & left(request("dqcode3"),ajlb_len3) &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where ajlb_code='" + request("dq")+"'"
  else'下移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where left(ajlb_code," & ajlb_len3 & ")='" & left(request("dqcode3"),ajlb_len3) &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where ajlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=301&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3 & "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3") & "&page4=" & cpage4)

elseif mode=408 then
  'delete 4类上移/下移
  opendb()
  if request("sort")="up" then'上移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where left(ajlb_code," & ajlb_len4 & ")='" & left(request("dqcode4"),ajlb_len4) &"' and mid(ajlb_code,"& (ajlb_len4+1) & "," & (ajlb_len5-ajlb_len4) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where ajlb_code='" + request("dq")+"'"
  else'下移
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where left(ajlb_code," & ajlb_len4 & ")='" & left(request("dqcode4"),ajlb_len4) &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len5-ajlb_len4) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where ajlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=401&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") &"&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") &"&page3=" & cpage3 & "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3") &"&page4=" & cpage4 & "&dqcode4="+request("dqcode4") & "&dqname4=" & request("dqname4") & "&page5=" & cpage5)
end if
%>    