<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if session("username")="" or instr(session("power"),",0,")=0 then
    Response.Redirect "notlogin.asp"
end if
%>

<!--#include file="fcommon.asp"-->
<!--#include file="dtp.asp"-->
<%
dim conn, mode, username, rs,rs1, sql, errmsg, founderror, s, t, i, fl, dq,odq,dq0, dq1,cpage1,cpage2,cpage3,cpage4,kpbm,st,dwxh,sfzs,dqcode1,dqcode2,dqcode3,dqcode4,dqname1,dqname2,dqname3,dqname4,sflc,qsc,xgsfdgls
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
  <title>��λ����</title>
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
        <b>��������</b>
      <%elseif mode>100 and mode<200 then %><!--��������-->
        <b>����[<%=request("dqname1")%>]-���оַ־��������</b>
      <%elseif mode>200 and mode<300 then %><!--С������-->
        <b>����[<%=request("dqname1")%>]-���оַ־�[<%=request("dqname2")%>]-�ɳ����������</b>
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
    <form method="POST" action="marea-1.asp?mode=2&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from unit where unit_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-1.asp?mode=3&page1=<%=cpage1%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-1.asp?mode=1&page1=<%=cpage1%>">����</a>]
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
            <td colspan="3">�����������Ȼ������OK��</td>
          <%else%>
            <td colspan="3">��༭������Ȼ������OK��</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=dq size=15 maxlength=<%=unit_len1%> class="smallInput" value='<%=request("dq")%>'>
            <%else%>
              <input name=dq size=15 maxlength=<%=unit_len1%> class="smallInput" value='<%=trim(left(rs("unit_code"),unit_len1))%>'>
            <%end if%>
            <font color=red>(*)</font>(��������ǰ<%=unit_len1%>λ,��<%=unit_len0-unit_len1%>λȫΪ0)
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=dq0 size=15 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=15 maxlength=30 class="smallInput" value='<%=trim(rs("unit_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ע-��λ&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=bbjzdwmc size=15 maxlength=30 class="smallInput" value=''>
            <%else%>
              <input name=bbjzdwmc size=15 maxlength=30 class="smallInput" value='<%=trim(rs("bbjzdwmc"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ע-����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=bbjzzg size=15 maxlength=30 class="smallInput" value=''>
            <%else%>
              <input name=bbjzzg size=15 maxlength=30 class="smallInput" value='<%=trim(rs("bbjzzg"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ע-����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=bbjzfh size=15 maxlength=30 class="smallInput" value=''>
            <%else%>
              <input name=bbjzfh size=15 maxlength=30 class="smallInput" value='<%=trim(rs("bbjzfh"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ע-�Ʊ�&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=bbjzzb size=15 maxlength=30 class="smallInput" value=''>
            <%else%>
              <input name=bbjzzb size=15 maxlength=30 class="smallInput" value='<%=trim(rs("bbjzzb"))%>'>
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
    <form method="POST" action="marea-1.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from unit where unit_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-1.asp?mode=103&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-1.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>">����</a>]
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
            <td colspan="3">�������շ�վ��Ȼ������OK��</td>
          <%else%>
            <td colspan="3">��༭�շ�վ��Ȼ������OK��</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շ�վ����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=dq size=15 maxlength=<%=unit_len3-unit_len2%> class="smallInput" value='<%=request("dq")%>'>(ǰ<%=unit_len1%>λΪ<%=left(request("dqcode1"),unit_len1)%>,�����<%=unit_len3-unit_len2%>λ)
            <%else%>
              <input name=dq size=15 maxlength=<%=unit_len3-unit_len2%> class="smallInput" value='<%=trim(mid(rs("unit_code"),unit_len1+1,unit_len2-unit_len1))%>'>(ǰ<%=unit_len1%>λΪ<%=left(request("dqcode1"),unit_len1)%>,�����<%=unit_len3-unit_len2%>λ)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շ�վ����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("unit_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ע-��λ&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=bbjzdwmc size=15 maxlength=30 class="smallInput" value=''>
            <%else%>
              <input name=bbjzdwmc size=15 maxlength=30 class="smallInput" value='<%=trim(rs("bbjzdwmc"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ע-����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=bbjzzg size=15 maxlength=30 class="smallInput" value=''>
            <%else%>
              <input name=bbjzzg size=15 maxlength=30 class="smallInput" value='<%=trim(rs("bbjzzg"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ע-����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=bbjzfh size=15 maxlength=30 class="smallInput" value=''>
            <%else%>
              <input name=bbjzfh size=15 maxlength=30 class="smallInput" value='<%=trim(rs("bbjzfh"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ע-�Ʊ�&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=bbjzzb size=15 maxlength=30 class="smallInput" value=''>
            <%else%>
              <input name=bbjzzb size=15 maxlength=30 class="smallInput" value='<%=trim(rs("bbjzzb"))%>'>
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

sub ShowInputForm201(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 202 then%>
    <form method="POST" action="marea-1.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    set rs1=server.createobject("adodb.recordset")
    rs.open "select * from unit where unit_code='" + request("odq") + "'", conn, 1, 1
    rs1.open "select * from sfdxx where wzlxcode='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-1.asp?mode=203&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-1.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>">����</a>]
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
            <td colspan="3">�������շѵ㣬Ȼ������OK��</td>
          <%else%>
            <td colspan="3">��༭�շѵ㣬Ȼ������OK��</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շѵ����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=dq size=15 maxlength=<%=unit_len2-unit_len1%> class="smallInput">(ǰ<%=unit_len2%>λΪ<%=left(request("dqcode2"),unit_len2)%>,�����<%=unit_len3%>λ)
            <%else%>
              <input name=dq size=15 maxlength=<%=unit_len2-unit_len1%> class="smallInput" value='<%=trim(mid(rs("unit_code"),unit_len2+1,unit_len3-unit_len2))%>'>(ǰ<%=unit_len2%>λΪ<%=left(request("dqcode2"),unit_len2)%>,�����<%=unit_len3%>λ)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շѵ�����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("unit_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��·����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=xlmc size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("xlmc")) then%>
                  <input name=xlmc size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("xlmc"))%>'>
                <%else%>
                  <input name=xlmc size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=xlmc size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��·���&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=xlbh size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("xlbh")) then%>
                  <input name=xlbh size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("xlbh"))%>'>
                <%else%>
                  <input name=xlbh size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=xlbh size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շ�·�����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=sfldqd size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("sfldqd")) then%>
                  <input name=sfldqd size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("sfldqd"))%>'>
                <%else%>
                  <input name=sfldqd size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=sfldqd size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շ�·��ֹ��&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=sfldzd size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("sfldzd")) then%>
                  <input name=sfldzd size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("sfldzd"))%>'>
                <%else%>
                  <input name=sfldzd size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=sfldzd size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�¸��շѵ�����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=xgsfdmc size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("xgsfdmc")) then%>
                  <input name=xgsfdmc size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("xgsfdmc"))%>'>
                <%else%>
                  <input name=xgsfdmc size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=xgsfdmc size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�¸��շѵ㹫��·&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=xgsfdgls size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("xgsfdgls")) then%>
                  <input name=xgsfdgls size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("xgsfdgls"))%>'>
                <%else%>
                  <input name=xgsfdgls size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=xgsfdgls size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շ�����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=sflx size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("sflx")) then%>
                  <input name=sflx size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("sflx"))%>'>
                <%else%>
                  <input name=sflx size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=sflx size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�����ȼ�&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=jsdj size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("jsdj")) then%>
                  <input name=jsdj size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("jsdj"))%>'>
                <%else%>
                  <input name=jsdj size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=jsdj size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�շ���ʼʱ��1&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzsfqshj size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzsfqshj")) then%>
                  <input name=pzsfqshj size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzsfqshj"))%>'>
                <%else%>
                  <input name=pzsfqshj size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzsfqshj size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�շѽ���ʱ��1&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzsfzshj size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzsfzshj")) then%>
                  <input name=pzsfzshj size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzsfzshj"))%>'>
                <%else%>
                  <input name=pzsfzshj size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzsfzshj size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�ĺ�1&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzwh size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzwh")) then%>
                  <input name=pzwh size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzwh"))%>'>
                <%else%>
                  <input name=pzwh size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzwh size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�շ���ʼʱ��2&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzsfqshj2 size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzsfqshj2")) then%>
                  <input name=pzsfqshj2 size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzsfqshj2"))%>'>
                <%else%>
                  <input name=pzsfqshj2 size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzsfqshj2 size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�շѽ���ʱ��2&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzsfzshj2 size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzsfzshj2")) then%>
                  <input name=pzsfzshj2 size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzsfzshj2"))%>'>
                <%else%>
                  <input name=pzsfzshj2 size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzsfzshj2 size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�ĺ�2&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzwh2 size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzwh2")) then%>
                  <input name=pzwh2 size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzwh2"))%>'>
                <%else%>
                  <input name=pzwh2 size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzwh2 size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�շ���ʼʱ��3&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzsfqshj3 size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzsfqshj3")) then%>
                  <input name=pzsfqshj3 size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzsfqshj3"))%>'>
                <%else%>
                  <input name=pzsfqshj3 size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzsfqshj3 size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�շѽ���ʱ��3&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzsfzshj3 size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzsfzshj3")) then%>
                  <input name=pzsfzshj3 size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzsfzshj3"))%>'>
                <%else%>
                  <input name=pzsfzshj3 size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzsfzshj3 size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��׼�ĺ�3&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=pzwh3 size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("pzwh3")) then%>
                  <input name=pzwh3 size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("pzwh3"))%>'>
                <%else%>
                  <input name=pzwh3 size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=pzwh3 size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շ����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=sflc size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("sflc")) then%>
                  <input name=sflc size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("sflc"))%>'>
                <%else%>
                  <input name=sflc size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=sflc size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=qsc size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("qsc")) then%>
                  <input name=qsc size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("qsc"))%>'>
                <%else%>
                  <input name=qsc size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=qsc size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�շ�����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=sfxz size=25 maxlength=30 class="smallInput" value=''>
            <%else
              if not rs1.eof then
                if not isnull(rs1("sfxz")) then%>
                  <input name=sfxz size=25 maxlength=30 class="smallInput" value='<%=trim(rs1("sfxz"))%>'>
                <%else%>
                  <input name=sfxz size=25 maxlength=30 class="smallInput" value=''>
                <%end if
              else%>
                <input name=sfxz size=25 maxlength=30 class="smallInput" value=''>                   
              <%end if
            end if%>
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
  if mode = 203 then
    rs.close
    rs1.close
    set rs1=nothing
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
  'response.write "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh"
  rs.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn, 1, 1
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
          [<a href="marea-1.asp?mode=2">���</a>]
          <%if cpage1 <> 1 then%>
            [<a href="marea-1.asp?mode=1&page1=<%=cstr(cpage1-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage1 <> rs.PageCount then%>
            [<a href="marea-1.asp?mode=1&page1=<%=cstr(cpage1+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage1 then%>
                <option selected value="marea-1.asp?mode=1&page1=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-1.asp?mode=1&page1=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>��������</td>
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
                <td align=center><%=trim(rs("unit_code"))%></td>
                <td align=center><%=trim(rs("unit_name"))%><a href="marea-1.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=trim(rs("unit_code"))%>&dqname1=<%=trim(rs("unit_name"))%>">��<font color="#FF0000">�շ�վ</font>��</a></td>
                <td align=center>
                  <a href="marea-1.asp?mode=3&page1=<%=cpage1%>&odq=<%=trim(rs("unit_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-1.asp?mode=4&page1=<%=cpage1%>&dq=<%=trim(rs("unit_code"))%>&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("unit_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-1.asp?mode=8&page1=<%=cpage1%>&dq=<%=trim(rs("unit_code"))%>&sort=up&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("unit_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-1.asp?mode=8&page1=<%=cpage1%>&dq=<%=trim(rs("unit_code"))%>&sort=down&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/down.gif" border=0></a>
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
          [<a href="marea-1.asp?mode=2">���</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">û�е�����¼</font></td></tr>
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
  'Response.Write("select * from unit where left(unit_code," & unit_len1 & ")='" & left(request("dqcode1"),unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh")
  rs.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(request("dqcode1"),unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage2>rs.pagecount then cpage2=1
    rs.AbsolutePage = cpage2
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">��<%=cstr(cpage2)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>���շ�վ��Ŀ</td>
          <td align="right">
          [<a href="marea-1.asp?mode=1&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">�����б�</a>]&nbsp;
          [<a href="marea-1.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">���</a>]
          <%if cpage2 <> 1 then%>
            [<a href="marea-1.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(cpage2-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage2 <> rs.PageCount then%>
            [<a href="marea-1.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(cpage2+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2" onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage2 then%>
                <option selected value="marea-1.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-1.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>�շ�վ����</td>
              <td width=40% align=center>�շ�վ����</td>
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
                <td align=center><%=trim(rs("unit_code"))%></td>
                <td align=center><%=trim(rs("unit_name"))%><a href="marea-1.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=trim(request("dqcode1"))%>&dqname1=<%=trim(request("dqname1"))%>&page2=<%=cpage2%>&dqcode2=<%=trim(rs("unit_code"))%>&dqname2=<%=trim(rs("unit_name"))%>">��<font color="#FF0000">�շѵ�</font>��</a></td>
                <td align=center>
                  <a href="marea-1.asp?mode=103&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&odq=<%=trim(rs("unit_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-1.asp?mode=104&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("unit_code"))%>&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("unit_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-1.asp?mode=108&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("unit_code"))%>&sort=up&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("unit_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-1.asp?mode=108&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("unit_code"))%>&sort=down&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/down.gif" border=0></a>
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
          [<a href="marea-1.asp?mode=1&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">�����б�</a>]&nbsp;
          [<a href="marea-1.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">���</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">û���շ�վ��¼</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=201 then
  'С����ʾ
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  rs.open "select * from unit where left(unit_code," & unit_len2 & ")='" & left(request("dqcode2"),unit_len2) &"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00' order by unit_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage3>rs.pagecount then cpage3=1
    rs.AbsolutePage = cpage3
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">��<%=cstr(cpage2)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>���շѵ���Ŀ</td>
          <td align="right">
          [<a href="marea-1.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>">�շ�վ�б�</a>]&nbsp;
          [<a href="marea-1.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>">���</a>]
          <%if cpage3 <> 1 then%>
            [<a href="marea-1.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cstr(cpage3-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage3 <> rs.PageCount then%>
            [<a href="marea-1.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cstr(cpage3+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage3 then%>
                <option selected value="marea-1.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-1.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>�շѵ����</td>
              <td width=40% align=center>�շѵ�����</td>
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
                <td align=center><%=trim(rs("unit_code"))%></td>
                <td align=center><%=trim(rs("unit_name"))%></td>
                <td align=center>
                  <a href="marea-1.asp?mode=203&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&odq=<%=trim(rs("unit_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-1.asp?mode=204&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("unit_code"))%>&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("unit_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-1.asp?mode=208&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("unit_code"))%>&sort=up&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("unit_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-1.asp?mode=208&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("unit_code"))%>&sort=down&dwxh=<%=trim(rs("unit_sxh"))%>"><img src="./images/down.gif" border=0></a>
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
          [<a href="marea-1.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>">�շ�վ�б�</a>]&nbsp;
          [<a href="marea-1.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>">���</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">û���շѵ��¼</font></td></tr>
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
    for i=len(dq) to unit_len1-1
      dq="0"+cstr(dq)
    next 
    for i=len(dq) to unit_len0-1
      dq=cstr(dq)+"0"
    next 
    dq0 = trim(request("dq0"))
    response.write dq
    if mode=2 then
      if dq = "" then
        ErrMsg="�������������"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select unit_name from unit where unit_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "���������ظ�"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn, 1, 1
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
        rs.open "select unit_name from unit where unit_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "���������ظ�"
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
        rs.open "unit", conn, 1, 3
        rs.addnew
        rs("unit_code")=dq
        rs("unit_name")=dq0
        rs("unit_sxh")=dwxh
        rs("bbjzdwmc")=request("bbjzdwmc")
        rs("bbjzzg")=request("bbjzzg")
        rs("bbjzfh")=request("bbjzfh")
        rs("bbjzzb")=request("bbjzzb")
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-1.asp?mode=1"
      else
        opendb()
        conn.Execute "update unit set unit_code='"+dq+"',unit_name='"+dq0+"',bbjzdwmc='"+request("bbjzdwmc")+"',bbjzzg='"+request("bbjzzg")+"',bbjzfh='"+request("bbjzfh")+"',bbjzzb='"+request("bbjzzb")+"' where unit_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-1.asp?mode=1&page1=" & cpage1
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
    for i=len(dq) to unit_len2-unit_len1-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode1"),unit_len1)+ dq
    for i=len(dq) to unit_len0-1
      dq=cstr(dq)+"0"
    next
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    if mode=102 then
      if dq = "" then
        ErrMsg="�������շ�վ����"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select unit_name from unit where unit_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "�շ�վ�����ظ�"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select unit_name from unit where left(unit_code," & unit_len1 & ")='" & left(request("dqcode1"),unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="�������շ�վ����"
        foundError=True
      end if
    else
      '���Ĺ����û����Ƿ����
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select unit_name from unit where unit_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "�շ�վ�����ظ�"
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
        rs.open "unit", conn, 1, 3
        rs.addnew
        rs("unit_code")=dq
        rs("unit_name")=dq0
        rs("unit_sxh")=dwxh
        rs("bbjzdwmc")=request("bbjzdwmc")
        rs("bbjzzg")=request("bbjzzg")
        rs("bbjzfh")=request("bbjzfh")
        rs("bbjzzb")=request("bbjzzb")
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-1.asp?mode=101&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")
      else
        opendb()
        conn.Execute "update unit set unit_code='"+dq+"',unit_name='"+dq0+"',bbjzdwmc='"+request("bbjzdwmc")+"',bbjzzg='"+request("bbjzzg")+"',bbjzfh='"+request("bbjzfh")+"',bbjzzb='"+request("bbjzzb")+"' where unit_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-1.asp?mode=101&page1="& cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2
      end if
    end if
  else
      ShowInputForm101 mode,""
  end if

elseif mode=202 or mode=203 then
  'С����Ӽ��޸�
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to unit_len3-unit_len2-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode2"),unit_len2)+ dq
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    if isnumeric(request("sflc")) then
      sflc=request("sflc")
    else
      sflc=0
    end if
    if isnumeric(request("qsc")) then
      qsc=request("qsc")
    else
      qsc=0
    end if
    if isnumeric(request("xgsfdgls")) then
      xgsfdgls=request("xgsfdgls")
    else
      xgsfdgls=0
    end if
    if mode=202 then
      if dq = "" then
        ErrMsg="�������շѵ����"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select unit_name from unit where unit_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "�շѵ�����ظ�"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select unit_name from unit where left(unit_code," & unit_len2 & ")='" & left(request("dqcode2"),unit_len2) &"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00' order by unit_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="�������շѵ�����"
        foundError=True
      end if
    else
      '���Ĺ����û����Ƿ����
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select unit_name from unit where unit_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "�շѵ�����ظ�"
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
        '�����
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "unit", conn, 1, 3
        rs.addnew
        rs("unit_code")=dq
        rs("unit_name")=dq0
        rs("unit_sxh")=dwxh
        rs.update
        rs.close
        '�����շѵ������Ϣ
        conn.execute "delete from sfdxx where wzlxcode='" + request("dq") + "'"
        rs.open "sfdxx", conn, 1, 3
        rs.addnew
        rs("xlmc")=request("xlmc")
        rs("xlbh")=request("xlbh")
        rs("sfldqd")=request("sfldqd")
        rs("sfldzd")=request("sfldzd")
        rs("xgsfdmc")=request("xgsfdmc")
        rs("xgsfdgls")=request("xgsfdgls")
        rs("zmcode")=request("dqcode2")
        rs("zm")=request("dqname2")
        rs("sflx")=request("sflx")
        rs("jsdj")=request("jsdj")
        rs("wzlxcode")=dq
        rs("wzlx")=dq0
        rs("pzsfqshj")=request("pzsfqshj")
        rs("pzsfzshj")=request("pzsfzshj")
        rs("pzwh")=request("pzwh")
        rs("pzsfqshj2")=request("pzsfqshj2")
        rs("pzsfzshj2")=request("pzsfzshj2")
        rs("pzwh2")=request("pzwh2")
        rs("pzsfqshj3")=request("pzsfqshj3")
        rs("pzsfzshj3")=request("pzsfzshj3")
        rs("pzwh3")=request("pzwh3")
        rs("sflc")=sflc
        rs("qsc")=qsc
        rs("sfxz")=request("sfxz")
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-1.asp?mode=201&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")& "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2")
      else
        opendb()
        conn.Execute "update unit set unit_code='"+dq+"',unit_name='"+dq0+"' where unit_code='"+request("odq")+"'"
        '�����շѵ������Ϣ
        conn.execute "delete from sfdxx where wzlxcode='" + request("odq") + "'"
        set rs=server.createobject("adodb.recordset")
        rs.open "sfdxx", conn, 1, 3
        rs.addnew
        rs("xlmc")=request("xlmc")
        rs("xlbh")=request("xlbh")
        rs("sfldqd")=request("sfldqd")
        rs("sfldzd")=request("sfldzd")
        rs("xgsfdmc")=request("xgsfdmc")
        rs("xgsfdgls")=xgsfdgls
        rs("zmcode")=request("dqcode2")
        rs("zm")=request("dqname2")
        rs("sflx")=request("sflx")
        rs("jsdj")=request("jsdj")
        rs("wzlxcode")=dq
        rs("wzlx")=dq0
        rs("pzsfqshj")=request("pzsfqshj")
        rs("pzsfzshj")=request("pzsfzshj")
        rs("pzwh")=request("pzwh")
        rs("pzsfqshj2")=request("pzsfqshj2")
        rs("pzsfzshj2")=request("pzsfzshj2")
        rs("pzwh2")=request("pzwh2")
        rs("pzsfqshj3")=request("pzsfqshj3")
        rs("pzsfzshj3")=request("pzsfzshj3")
        rs("pzwh3")=request("pzwh3")
        rs("sflc")=sflc
        rs("qsc")=qsc
        rs("sfxz")=request("sfxz")
        rs.update
        rs.close
        set rs=nothing
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-1.asp?mode=201&page1="& cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3
      end if
    end if
  else
      ShowInputForm201 mode,""
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
      ���Ҫɾ��������<%=request("dq")%>����
      <br><br>
      [<a href="marea-1.asp?mode=7&page1=<%=cpage1%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-1.asp?mode=1&page1=<%=cpage1%>">����</a>]
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
      ���Ҫɾ���շ�վ��<%=request("dq")%>����
      <br><br>
      [<a href="marea-1.asp?mode=107&page1=<%=cpage1%>&page2=<%=cpage2%>&dq=<%=request("dq")%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&dwxh=<%=request("dwxh")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-1.asp?mode=101&page1=<%=cpage1%>&page2=<%=cpage2%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">����</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=204 then
  'С��ɾ��ȷ��
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      ���Ҫɾ���շѵ㡰<%=request("dq")%>����
      <br><br>
      [<a href="marea-1.asp?mode=207&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-1.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>">����</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=7 then
  '����delete
  opendb()
  conn.execute "delete from unit where unit_code like '" + left(request("dq"),unit_len1)+"%'"'���������༰�����������С��
  conn.execute "delete from sfdxx where wzlxcode like '" + left(request("dq"),unit_len1) + "%'"'����������շѵ������Ϣ
  conn.execute "update unit set unit_sxh=unit_sxh-1 where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' and unit_sxh>=" & request("dwxh")' �����˳�����ǰ��
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-1.asp?mode=1&page1=" & cpage1)

elseif mode=107 then
  '����delete
  opendb()
  conn.execute "delete from unit where unit_code like '" + left(request("dq"),unit_len2)+"%'"'����������༰������С��
  conn.execute "delete from sfdxx where wzlxcode like '" + left(request("dq"),unit_len2) + "%'"'����������շѵ������Ϣ
  conn.execute "update unit set unit_sxh=unit_sxh-1 where left(unit_code," & unit_len1 & ")='" & left(request("dqcode1"),unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' and unit_sxh>=" & request("dwxh")' �����˳�����ǰ��
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-1.asp?mode=101&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2)  

elseif mode=207 then
  'С��delete
  opendb()
  conn.execute "delete from unit where unit_code like '" + left(request("dq"),unit_len3)+"%'"'�������С��
  conn.execute "delete from sfdxx where wzlxcode like '" + left(request("dq"),unit_len3) + "%'"'����������շѵ������Ϣ
  conn.execute "update unit set unit_sxh=unit_sxh-1 where left(unit_code," & unit_len2 & ")='" & left(request("dqcode2"),unit_len2) &"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00' and unit_sxh>=" & request("dwxh")' �����˳�����ǰ��
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-1.asp?mode=201&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") &"&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3)

elseif mode=8 then
  'delete ��������/����
  opendb()
  if request("sort")="up" then'����
    conn.execute "update unit set unit_sxh=unit_sxh+1 where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' and unit_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update unit set unit_sxh=unit_sxh-1 where unit_code='" + request("dq")+"'"
  else'����
    conn.execute "update unit set unit_sxh=unit_sxh-1 where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' and unit_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update unit set unit_sxh=unit_sxh+1 where unit_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-1.asp?mode=1&page1=" & cpage1)

elseif mode=108 then
  'delete ��������/����
  opendb()
  if request("sort")="up" then'����
    conn.execute "update unit set unit_sxh=unit_sxh+1 where left(unit_code," & unit_len1 & ")='" & left(request("dqcode1"),unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' and unit_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update unit set unit_sxh=unit_sxh-1 where unit_code='" + request("dq")+"'"
  else'����
    conn.execute "update unit set unit_sxh=unit_sxh-1 where left(unit_code," & unit_len1 & ")='" & left(request("dqcode1"),unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' and unit_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update unit set unit_sxh=unit_sxh+1 where unit_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-1.asp?mode=101&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2)

elseif mode=208 then
  'delete С������/����
  opendb()
  if request("sort")="up" then'����
    conn.execute "update unit set unit_sxh=unit_sxh+1 where left(unit_code," & unit_len2 & ")='" & left(request("dqcode2"),unit_len2) &"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00' and unit_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update unit set unit_sxh=unit_sxh-1 where unit_code='" + request("dq")+"'"
  else'����
    conn.execute "update unit set unit_sxh=unit_sxh-1 where left(unit_code," & unit_len2 & ")='" & left(request("dqcode2"),unit_len2) &"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00' and unit_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update unit set unit_sxh=unit_sxh+1 where unit_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-1.asp?mode=201&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&page3=" & cpage3)
end if
%>    