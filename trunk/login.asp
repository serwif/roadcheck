<%@ LANGUAGE="VBSCRIPT"%>
<%option explicit%>

<!--#include file="fcommon.asp"-->

<%
dim conn_system
dim cpage
dim username
dim password
dim rs,sql
dim FoundError
dim ErrMsg

sub opendb()
  'open system
  set conn_system=server.createobject("ADODB.CONNECTION")
  on error resume next
  err.clear
  conn_system.open sysconstr
  if err.number<>0 then
    err.clear
    Response.Redirect "error.asp?errid=1"
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
  <title>��¼</title>
  <link rel="stylesheet" type="text/css" href="./main.css">
  </head>

<script language="javascript">

function register()
{
   //location.href="muser.asp?mode=2&register=1";
   location.href="main11.asp?register=1";
   return false;       
}

function visitor()
{
   //location.href="searchmryl.asp?mode=1&visitor=1";
   location.href="main10.asp?visitor=1";
   return false;   
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

function showinputform0(errmsg)
  'on error resume next
  if errmsg="" then errmsg="&nbsp;"
  showchead()
  Randomize
%>
  <form action="login.asp" method="POST" name="input0">
  <table  width=100% height=100% border=0 align=center><tr><td valign="middle">
  <table background="./images/loginbk.jpg" width=100% height="80" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td></td>
    </tr>
  </table>
  <table  width=100% height="80" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr><td>
    <table width="400" border="0" cellspacing="0" cellpadding="0" align=right>
      <tr><td><font color=red><b><%=errmsg%></b></font><br></td></tr>
      <tr>
        <td>
          �û���<input class="smallInput" name="username" size="12" maxlength="6" value="<%=request("username")%>">
          ���룺<input type=password class="smallInput" name="password" size="12" maxlength="10" value="<%=request("password")%>">
          &nbsp;<input class="buttonface" type="submit" value=" ��¼ " id=submit1 name=submit1>
          &nbsp;<input class="buttonface" type="button" value=" ע�� " id=submit2 name=submit2 onclick="register()">         
        </td>
      </tr>
    </table>
  </td></tr>
  </table>
  <table width="100%">
    <tr>
      <td><font color=696969>��ѷֱ���1024X768��</font></td>
      <td align="right"><font color=696969>�����н�ͨ��--������¡���繤�����޹�˾</font></td>
    </tr>
  </table>
  </td></tr></table>
  </form>
<%
  showctail()
end function

username=trim(request("username"))
password=request("password")

session("userlogin")=false

if username<>"" or password<>"" then
  FoundError=false
  if UserName="" then
    founderror=true
    errmsg = "�������û���"
  end if
  if password="" then
    founderror=true
    if errmsg<>"" then
    	errmsg=errmsg+"������"
    else
      errmsg="����������"
    end if
  end if
  if founderror then
    ShowInputForm0 errmsg
  else
      '�鿴�û�
      opendb()
      set rs=server.createobject("adodb.recordset")
      rs.open "select name,unit_code,unit_name,power,skin from userinfo where username='"&username&"' and password='"&password&"'",conn_system,1,1
      'rs.open "select power,skin from user where name='"&username&"' and password='"&password&"'",conn_system,1,1
      if rs.recordcount=0 then
        rs.close
        set rs=nothing
        closedb()
        ShowInputForm0 "��¼ʧ�ܣ������û���������"
      else
        session("username")=trim(username)
        session("password")=password
        session("power")=trim(rs("power"))
        session("skin")=trim(rs("skin"))
	session("name")=trim(rs("name"))
	session("unit_code")=trim(rs("unit_code"))
	session("unit_name")=trim(rs("unit_name"))
        session("area_code")=left(rs("unit_code"),unit_len1)
       	'sql="select * from unit where unit_code='"&left(rs("unit_code"),6)&"'"
    	if session("skin")="" then 
          session("skin")="green"
        end if
        if username="admin" then
          session("menu")="7"
        else
          session("menu")="2"
        end if
        'Response.Cookies("skin") = session("skin")
        'Response.Cookies("skin").expires = date+30
        rs.close
	'session("unit_name")=GetUnitName(session("unit_code"))
        'session("unit_name")=""
	'rs.Open "select * from systemconfig",conn_system,1,1
	'if not rs.EOF then
	'  if not isnull(rs("area_code")) then 
	'    session("area_code")=rs("area_code")
	'  end if
	'end if
	'rs.Close 
        '�����¼��־
        if username="admin" then
          conn_system.execute("insert into olog (shj,username,czms,bz) values ('"&now()&"','"&username&"','ϵͳ����Ա��¼','DL')")
        else
          conn_system.execute("insert into olog (shj,username,czms,bz) values ('"&now()&"','"&username&"','��ͨ�ڲ��û���¼','DL')")
        end if
        set rs=nothing
        closedb()
        response.redirect("main"&session("menu")&".asp")
      end if
  end if
else
  ShowInputForm0("")
end if

function GetUnitName(s)
  sql="select * from unit where unit_code='"&left(s,6)&"'"
  set rs=conn_system.execute(sql)
  if rs.eof and rs.bof then
    rs.close:set rs=nothing
    getunitname="":exit function
  else
    getunitname=rs("unit_name")
  end if
  rs.close
end function
%>