<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>
<!--#include file="./fcommon.asp"--> 
<%
if session("username")="" then
  Response.Redirect("notlogin.asp")
end if

dim conn_system, rs
dim username, opassword, password, vpassword
dim FoundError, ErrMsg, fl

sub opendb()
  set conn_system=server.createobject("ADODB.CONNECTION")
  if err.number<>0 then
    err.clear
    Response.Redirect "error.asp?errid=1"
  else
    conn_system.open sysconstr
    if err then
      err.clear
      Response.Redirect "error.asp?errid=1"
    end if
  end if
end sub

sub closedb()
  conn_system.Close
  set conn_system=nothing
end sub

sub showchead()
%>
  <html>
  <head>
  <meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312">
  <title>�û������޸�</title>
  <link rel="stylesheet" type="text/css" href="main.css">
  </head>
  <body>
  <%noRightClick()
end sub

sub showctail()
%>
  </body>
  </html>
<%
end sub

sub ShowInputForm(ErrMsg)
  on error resume next
  showchead()
%>
  <table width="530" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>�޸�����</b></td>
    </tr>
    <tr><td align="center">

    <form method="POST" action="changepwd.asp" name="newuser">
      <table width="500" border="0" cellspacing="1" cellpadding="3" align="center">
        <tr><td colspan="2">&nbsp;</td></tr>
        <tr><td colspan="2" align="center"><%=Errmsg%></td></tr>
        <tr><td colspan="2"><hr width="100%" noshade align="left" size="1"></td></tr>
        <tr>
          <td width="100" bgcolor="#eeeeee" align="right">������&nbsp;</td>
          <td width="400">
            <input type=password name=opassword size=25 maxlength=20 class="smallInput" value='<%=request("opassword")%>'><font color="#FF0000">��������д��</font>
          </td>
        </tr>
        <tr>
          <td width="100" bgcolor="#eeeeee" align="right">������&nbsp;</td>
          <td width="400">
            <input type=password name=password size=25 maxlength=20 class="smallInput" value='<%=request("password")%>'><font color="#FF0000">��������д��</font>
            <br>�20λ�����������ַ��������鲻Ҫ�����յ����ֻ򵥴���Ϊ���롣
          </td>
        </tr>
        <tr>
          <td width="100" bgcolor="#eeeeee" align="right">У������&nbsp;</td>
          <td width="400">
            <input type=password name=vpassword size=25 maxlength=20 class="smallInput" value='<%=request("vpassword")%>'><font color="#FF0000">��������д��</font>
          </td>
        </tr>
        <tr><td colspan="2"><hr width="100%" noshade align="left" size="1"></td></tr>
        <tr><td colspan="2" align="center">
          <input class="buttonface" type="submit" value=" �� �� " id=submit1 name=submit1>
        </td></tr>
        </table>
      </form>
    </td></tr>
<%
  showctail()
end sub

  if request("opassword")<>"" or request("password")<>"" or request("vpassword")<>"" then
    'get input information
    username = session("username")
    opassword=request("opassword")
    password=request("password")
    vpassword=request("vpassword")
    if username="" or opassword="" then
      ErrMsg="�������û�����������"
      foundError=True
    else
      opendb()
      set rs=server.createobject("adodb.recordset")
      rs.open "select username from userinfo where username='" + username + "' and password='"+opassword+"'", conn_system, 1, 1
      if rs.recordcount=0 then
        ErrMsg = "���������"
        FoundError = True
      end if
      rs.close
      set rs=nothing
      closedb()
    end if
    if password = "" then
      if ErrMsg <> "" then
        ErrMsg = ErrMsg + "������"
      else
        ErrMsg = "����������������"
        foundError=True
      end if
    else
      if vpassword="" then
        if ErrMsg <> "" then
          ErrMsg = ErrMsg + "��У������"
        else
          ErrMsg = "����������У������"
          foundError=True
        end if
      else
        if password<>vpassword then
          if ErrMsg <> "" then
            ErrMsg = ErrMsg + "��������������벻һ��"
          else
            ErrMsg = "������������벻һ��"
            foundError=True
          end if
        end if
      end if
    end if
    if FoundError then
      ShowInputForm ErrMsg
    else
      on error resume next
      err.clear
      opendb()
      conn_system.execute "update userinfo set password='"+password+"' where username='"+username+"'"
      closedb()
      showchead()
%>
  <table width="90%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>�޸�����</b></td>
    </tr>
    <tr><td align="center"><br>
      �����Ѿ��޸ĳɹ����´ε�¼ʱ��ǵ��������롣����<br><br>
    </td></tr>
  </table>
<%
      showctail()
    end if
  else
    ShowInputForm "����д�����Ϣ"
  end if
%>