<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if session("username")="" or instr(session("power"),",0,")=0 then
    Response.Redirect "notlogin.asp"
end if
%>

<!--#include file="fcommon.asp"-->

<%
dim conn, mode, username, rs, sql, errmsg, founderror, s, t, i, fl, dq,odq,dq0, dq1,cpage1,cpage2,cpage3,cpage4,kpbm,st,dwxh,sfzs,dqcode1,dqcode2,dqcode3,dqcode4,dqname1,dqname2,dqname3,dqname4
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
  <title>������</title>
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
      <%if mode<100 then %><!--��������-->
        <b>�����������</b>
      <%elseif mode>100 and mode<200 then %><!--��������-->
        <b>����[<%=request("dqname1")%>]-С���������</b>
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
        [<a href="marea-2.asp?mode=1&page1=<%=cpage1%>">����</a>]
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
            <td colspan="3">��������࣬Ȼ������OK��</td>
          <%else%>
            <td colspan="3">��༭���࣬Ȼ������OK��</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=dq size=15 maxlength=<%=ajlb_len1%> class="smallInput" value='<%=request("dq")%>'>
            <%else%>
              <input name=dq size=15 maxlength=<%=ajlb_len1%> class="smallInput" value='<%=trim(left(rs("ajlb_code"),ajlb_len1))%>'>
            <%end if%>
            <font color=red>(*)</font>(��������ǰ<%=ajlb_len1%>λ,��<%=ajlb_len0-ajlb_len1%>λȫΪ0)
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��������&nbsp;</td>
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
          <td colspan="3"><input class="buttonface" type="submit" value=" ȷ �� " id=submit1 name=submit1></td>
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
        [<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>">����</a>]
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
            <td colspan="3">������С�࣬Ȼ������OK��</td>
          <%else%>
            <td colspan="3">��༭С�࣬Ȼ������OK��</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>С�����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=dq size=15 maxlength=<%=ajlb_len3-ajlb_len2%> class="smallInput" value='<%=request("dq")%>'>(ǰ<%=ajlb_len1%>λΪ<%=left(request("dqcode1"),ajlb_len1)%>,�����<%=ajlb_len3-ajlb_len2%>λ)
            <%else%>
              <input name=dq size=15 maxlength=<%=ajlb_len3-ajlb_len2%> class="smallInput" value='<%=trim(mid(rs("ajlb_code"),ajlb_len1+1,ajlb_len2-ajlb_len1))%>'>(ǰ<%=ajlb_len1%>λΪ<%=left(request("dqcode1"),ajlb_len1)%>,�����<%=ajlb_len3-ajlb_len2%>λ)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>С������&nbsp;</td>
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
          <td colspan="3"><input class="buttonface" type="submit" value=" ȷ �� " id=submit1 name=submit1></td>
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

if mode=1 then
  '������ʾ 
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
          <td valign="bottom">��<%=cstr(cpage1)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>��������Ŀ</td>
          <td align="right">
          [<a href="marea-2.asp?mode=2">���</a>]
          <%if cpage1 <> 1 then%>
            [<a href="marea-2.asp?mode=1&page1=<%=cstr(cpage1-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage1 <> rs.PageCount then%>
            [<a href="marea-2.asp?mode=1&page1=<%=cstr(cpage1+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage1 then%>
                <option selected value="marea-2.asp?mode=1&page1=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-2.asp?mode=1&page1=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>�������</td>
              <td width=40% align=center>��������</td>
              <td width=50% align=center>����</td>
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
                <td align=center><%=trim(rs("ajlb_name"))%><a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=trim(rs("ajlb_code"))%>&dqname1=<%=trim(rs("ajlb_name"))%>">��<font color="#FF0000">����</font>��</a></td>
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
          [<a href="marea-2.asp?mode=2">���</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">û�д����¼</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()
elseif mode=101 then
  '������ʾ
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
          <td valign="bottom">��<%=cstr(cpage2)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>��С����Ŀ</td>
          <td align="right">
          [<a href="marea-2.asp?mode=1&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">�����б�</a>]&nbsp;
          [<a href="marea-2.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">���</a>]
          <%if cpage2 <> 1 then%>
            [<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(cpage2-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage2 <> rs.PageCount then%>
            [<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(cpage2+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2" onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage2 then%>
                <option selected value="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-2.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>С�����</td>
              <td width=40% align=center>С������</td>
              <td width=50% align=center>����</td>
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
          [<a href="marea-2.asp?mode=1&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">�����б�</a>]&nbsp;
          [<a href="marea-2.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">���</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">û��С���¼</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=2 or mode=3 then
  '������Ӽ��޸�
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
        ErrMsg="������������"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "��������ظ�"
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
        ErrMsg="�������������"
        foundError=True
      end if
    else
      '���Ĺ����û����Ƿ����
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "��������ظ�"
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
        '�����
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
  '������Ӽ��޸�
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
        ErrMsg="������С�����"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "С������ظ�"
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
        ErrMsg="������С������"
        foundError=True
      end if
    else
      '���Ĺ����û����Ƿ����
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select ajlb_name from ajlb where ajlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "С������ظ�"
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
        '�����
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

elseif mode=4 then
  '����ɾ��ȷ��
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      ���Ҫɾ�����ࡰ<%=request("dq")%>����
      <br><br>
      [<a href="marea-2.asp?mode=7&page1=<%=cpage1%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-2.asp?mode=1&page1=<%=cpage1%>">����</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=104 then
  '����ɾ��ȷ��
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      ���Ҫɾ��С�ࡰ<%=request("dq")%>����
      <br><br>
      [<a href="marea-2.asp?mode=107&page1=<%=cpage1%>&page2=<%=cpage2%>&dq=<%=request("dq")%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&dwxh=<%=request("dwxh")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-2.asp?mode=101&page1=<%=cpage1%>&page2=<%=cpage2%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">����</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=7 then
  '����delete
  opendb()
  conn.execute "delete from ajlb where ajlb_code like '" + left(request("dq"),ajlb_len1)+"%'"'���������༰�����������С��
  conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' and ajlb_sxh>=" & request("dwxh")' �����˳�����ǰ��
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=1&page1=" & cpage1)

elseif mode=107 then
  '����delete
  opendb()
  conn.execute "delete from ajlb where ajlb_code like'" + left(request("dq"),ajlb_len2)+"%'"'����������༰������С��
  conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' and ajlb_sxh>=" & request("dwxh")' �����˳�����ǰ��
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=101&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2)  

elseif mode=8 then
  'delete ��������/����
  opendb()
  if request("sort")="up" then'����
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' and ajlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where ajlb_code='" + request("dq")+"'"
  else'����
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' and ajlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where ajlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=1&page1=" & cpage1)

elseif mode=108 then
  'delete ��������/����
  opendb()
  if request("sort")="up" then'����
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where ajlb_code='" + request("dq")+"'"
  else'����
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh-1 where left(ajlb_code," & ajlb_len1 & ")='" & left(request("dqcode1"),ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' and ajlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update ajlb set ajlb_sxh=ajlb_sxh+1 where ajlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-2.asp?mode=101&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2)
end if
%>    