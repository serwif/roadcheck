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
  <title>用户密码修改</title>
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
      <td align="center"><b>修改密码</b></td>
    </tr>
    <tr><td align="center">

    <form method="POST" action="changepwd.asp" name="newuser">
      <table width="500" border="0" cellspacing="1" cellpadding="3" align="center">
        <tr><td colspan="2">&nbsp;</td></tr>
        <tr><td colspan="2" align="center"><%=Errmsg%></td></tr>
        <tr><td colspan="2"><hr width="100%" noshade align="left" size="1"></td></tr>
        <tr>
          <td width="100" bgcolor="#eeeeee" align="right">旧密码&nbsp;</td>
          <td width="400">
            <input type=password name=opassword size=25 maxlength=20 class="smallInput" value='<%=request("opassword")%>'><font color="#FF0000">（必须填写）</font>
          </td>
        </tr>
        <tr>
          <td width="100" bgcolor="#eeeeee" align="right">新密码&nbsp;</td>
          <td width="400">
            <input type=password name=password size=25 maxlength=20 class="smallInput" value='<%=request("password")%>'><font color="#FF0000">（必须填写）</font>
            <br>最长20位，可用任意字符，但建议不要用生日等数字或单词作为密码。
          </td>
        </tr>
        <tr>
          <td width="100" bgcolor="#eeeeee" align="right">校验密码&nbsp;</td>
          <td width="400">
            <input type=password name=vpassword size=25 maxlength=20 class="smallInput" value='<%=request("vpassword")%>'><font color="#FF0000">（必须填写）</font>
          </td>
        </tr>
        <tr><td colspan="2"><hr width="100%" noshade align="left" size="1"></td></tr>
        <tr><td colspan="2" align="center">
          <input class="buttonface" type="submit" value=" 提 交 " id=submit1 name=submit1>
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
      ErrMsg="请输入用户名及旧密码"
      foundError=True
    else
      opendb()
      set rs=server.createobject("adodb.recordset")
      rs.open "select username from userinfo where username='" + username + "' and password='"+opassword+"'", conn_system, 1, 1
      if rs.recordcount=0 then
        ErrMsg = "旧密码错误！"
        FoundError = True
      end if
      rs.close
      set rs=nothing
      closedb()
    end if
    if password = "" then
      if ErrMsg <> "" then
        ErrMsg = ErrMsg + "，密码"
      else
        ErrMsg = "请输入您的新密码"
        foundError=True
      end if
    else
      if vpassword="" then
        if ErrMsg <> "" then
          ErrMsg = ErrMsg + "，校验密码"
        else
          ErrMsg = "请输入您的校验密码"
          foundError=True
        end if
      else
        if password<>vpassword then
          if ErrMsg <> "" then
            ErrMsg = ErrMsg + "，两次输入的密码不一致"
          else
            ErrMsg = "两次输入的密码不一致"
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
      <td align="center"><b>修改密码</b></td>
    </tr>
    <tr><td align="center"><br>
      密码已经修改成功，下次登录时请记得用新密码。：）<br><br>
    </td></tr>
  </table>
<%
      showctail()
    end if
  else
    ShowInputForm "请填写相关信息"
  end if
%>