<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if request("register")=1 then
elseif session("username")="" or instr(session("power"),",0,")=0 then
    Response.Redirect "notlogin.asp?title=�û�����"
end if
%>

<!--#include file="fcommon.asp"-->
<!--#include file="dtp.asp"-->

<%
dim conn_system, mode,register, username, rs,rsMX,rs1, sql, errmsg, founderror, s, t, i, fl, memname,cpage,truename,password,workphone,handset,familialphone,FRMunit,FRMbusiness,workshj,FRMdw,FRMpcs
dim FRMcsrq,FRMrjrq,FRMwhcd
dim unit_code,unit_name

if not isempty(request("mode")) then
    mode = clng(request("mode"))
else
    mode=1
end if
if not isempty(request("register")) then
    register = clng(request("register"))
else
    register=0
end if
if not isempty(request("username")) then
    username = request("username")
else
    username = ""
end if

if not isempty(request("FRMdw"))  then
    FRMdw =  trim(request("FRMdw"))
else
    FRMdw=""
end if

if not isempty(request("FRMpcs"))  then
    FRMpcs =  trim(request("FRMpcs"))
else
    FRMpcs=""
end if
if not isempty(request("unit_code")) then
    unit_code = request("unit_code")
else
    unit_code = ""
end if

sub opendb()
  set conn_system=server.createobject("ADODB.CONNECTION")
  conn_system.open sysconstr
end sub

sub closedb()
  conn_system.Close
  set conn_system=nothing
end sub

sub showchead()
%>
  <html>
  <head>
  <title><%
  if register=1 then %>
�û�ע��
  <%else%>
�û�����
  <%end if%>
</title>
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
  
  function Getseconditem(i,j)
  {//������С���б�
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
     location.href="muser.asp?mode=2&unit_code="+unit_code+"&username="+document.input1.username1.value+"&register="+document.input1.register1.value;
   else
     {if(i==103)
        location.href="muser.asp?mode=3&unit_code="+unit_code+"&username="+document.input1.username1.value+"&register="+document.input1.register1.value;
      else
        {if(i==202)
           {
           //alert("muser.asp?mode=12&unit_code="+unit_code+"&username="+document.input1.username1.value+"&register="+document.input1.register1.value);
           location.href="muser.asp?mode=12&unit_code="+unit_code+"&username="+document.input1.username1.value+"&register="+document.input1.register1.value;
           }
         else
           {if(i==203)
              {
              //alert("muser.asp?mode=13&unit_code="+unit_code+"&username="+document.input1.username1.value+"&register="+document.input1.register1.value);
              location.href="muser.asp?mode=13&unit_code="+unit_code+"&username="+document.input1.username1.value+"&register="+document.input1.register1.value;
              }
            else    
              location.href="muser.asp?mode=1&unit_code="+unit_code;             
           }
        } 
      }
   return false;
  }
  //-->
  </script>

  <body>
  <%noRightClick()%>
  <table width="90%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28"><td align="center">
      <b>
      <%if register=1 then %>
�û�ע��
      <%else%>
�û�����
      <%end if%></b>
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

sub drawPowerCheck(s,t)
  if instr(s,","+cstr(t)+",")<>0 then
    response.write "<img src='./images/checked.gif' border='0'>"
  else
    response.write "<img src='./images/unchecked.gif' border='0'>"
  end if
end sub

sub isPowerCheck(s,t)
  if instr(s,","+cstr(t)+",")<>0 then
    response.write "checked"
  end if
end sub

sub ShowInputForm1(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 2 then
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn_system, 1, 1%>
    <form method="POST" action="muser.asp?mode=2&username=<%=username%>&register=<%=register%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn_system, 1, 1
    rs.open "select * from userinfo where username='" + username + "'", conn_system, 1, 1%>
    <form method="POST" action="muser.asp?mode=3&username=<%=username%>&register=<%=register%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        <%if register=1 then%>
        <%else%>
          <%if username="" then%>
            [<a href="muser.asp?mode=1">����</a>]
          <%else%>
            [<a href="muser.asp?mode=8&username=<%=username%>">����</a>]
          <%end if%>
        <%end if%>
      </td>
    </tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
    <tr><td>
      <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
        <tr>
            <%if Trim(ErrMsg) <> "" then%>
              <%=errmsg%>
            <%else%>
              <% if mode = 2 then%>
                �������û���Ϣ��Ȼ������OK��
              <%else%>
                ��༭�û���Ϣ��Ȼ������OK��
                <input name="odq" type="hidden" value="<%=request("odq")%>">
              <%end if%>
            <%end if%>
            <input name="username1" type="hidden" value="<%=username%>">
            <input name="register1" type="hidden" value="<%=register%>">
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�û���<font color="#FF0000">*</font>&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=memname size=15 maxlength=6 class="smallInput" value='<%=request("memname")%>' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
            <%else
              if rs("username")<>"admin" then%>
                <input name=memname size=15 maxlength=6 class="smallInput" value='<%=rs("username")%>' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
              <%else%>
                admin
              <%end if%>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��ʵ����<font color="#FF0000">*</font>&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=truename size=15 maxlength=10 class="smallInput" value='<%=request("truename")%>'>
            <%else
              if rs("username")<>"admin" then%>
                <input name=truename size=15 maxlength=10 class="smallInput" value='<%=rs("name")%>'>
              <%else%>
                admin
              <%end if%>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" class="smallInput" align=right nowrap>����<font color="#FF0000">*</font>&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=password size=15 type="password" maxlength=10 class="smallInput" value='<%=request("password")%>'>
            <%else%>
              <input name=password size=15 type="password" maxlength=10 class="smallInput" value='<%=rs("password")%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" class="smallInput" align=right nowrap>���ڵ�λ<font color="#FF0000">*</font>&nbsp;</td>
          <td align=left colspan=2>
	    ����
            <%if mode=2 then%> 
              <select name="unit_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(102,1)">
            <%else%>
              <select name="unit_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(103,1)">
            <%end if%>
            <%while not rs1.EOF 
              if trim(unit_code)="" then unit_code=trim(rs1("unit_code"))%>
              <option value="<%=trim(rs1("unit_code"))%>"<%if left(unit_code,unit_len1)=left(rs1("unit_code"),unit_len1) then %> selected <% end if %>><%=trim(rs1("unit_name"))%></option>
              <%rs1.MoveNext 
            WEND%>
            </select>
            <br>�շ�վ
            <%if mode=2 then%> 
              <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(102,2)">
            <%else%>
              <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(103,2)">
            <%end if%>
            <option value="" <%if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then %> selected <% end if %>></option>
            <%rsMX.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn_system, 1, 1
            while not rsMX.EOF
              'if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then unit_code=trim(rsMX("unit_code"))%>
              <option value="<%=trim(rsMX("unit_code"))%>"<%if left(unit_code,unit_len2)=left(rsMX("unit_code"),unit_len2) then %> selected <% end if %>><%=trim(rsMX("unit_name"))%></option>
              <%rsMX.MoveNext 
            WEND
            rsMX.close%>
            </select>
          </td>
        </tr>
        <%if register=1 then%>
        <%else%>
          <tr>
            <td bgcolor="#eeeeee" align=right nowrap>Ȩ��<font color="#FF0000">*</font>&nbsp;</td>
            <td align=left valign=center colspan=2>
              <%if mode=3 then
                if rs("username")<>"admin" then%>
                  <input type=checkbox name=power value='1' <%isPowerCheck rs("power"),1%>>���ݵǼ�<br>
                  <input type=checkbox name=power value='2' <%isPowerCheck rs("power"),2%>>���ݱ��<br>
                  <input type=checkbox name=power value='3' <%isPowerCheck rs("power"),3%>>��ѯͳ��<br>
                <%end if%>
              <%else%>
                <input type=checkbox name=power value='0'>ϵͳ����<br>
                <input type=checkbox name=power value='1'>���ݵǼ�<br>
                <input type=checkbox name=power value='2'>���ݱ��<br>
                <input type=checkbox name=power value='3'>��ѯͳ��<br>
              <%end if%>
            </td>
          </tr>
        <%end if%>
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
  else
    closedb()
  end if
  showctail
end sub

sub ShowInputForm11(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 12 then
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn_system, 1, 1%>
    <form method="POST" action="muser.asp?mode=12&username=<%=username%>&register=<%=register%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn_system, 1, 1
    rs.open "select * from userinfo where username='" + username + "'", conn_system, 1, 1%>
    <form method="POST" action="muser.asp?mode=13&username=<%=username%>&register=<%=register%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        <%if register=1 then%>
        <%else%>
        <%if username="" then%>
          [<a href="muser.asp?mode=1">����</a>]
        <%else%>
          [<a href="muser.asp?mode=18&username=<%=username%>">����</a>]
        <%end if%>
        <%end if%>
      </td>
    </tr>
    <tr><td><hr noshade size="1" width="100%"></td></tr>
    <tr><td>
      <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
        <tr>
          <td colspan="3">
            <%if Trim(ErrMsg) <> "" then%>
              <%=errmsg%>
            <%else%>
              <% if mode = 12 then%>
                �������û���Ϣ��Ȼ������OK��
              <%else%>
                ��༭�û���Ϣ��Ȼ������OK��
                <input name="odq" type="hidden" value="<%=request("odq")%>">
              <%end if%>
            <%end if%>
            <input name="username1" type="hidden" value="<%=username%>">
            <input name="register1" type="hidden" value="<%=register%>">
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�û���<font color="#FF0000">*</font>&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=12 then%>
              <input name=memname size=15 maxlength=6 class="smallInput" value='<%=request("memname")%>' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
            <%else
              if rs("username")<>"admin" then%>
                <input name=memname size=15 maxlength=6 class="smallInput" value='<%=rs("username")%>' onkeydown="if((window.event.keyCode<48 || window.event.keyCode>57) && window.event.keyCode!=190 && window.event.keyCode!=8 || window.event.shiftKey){window.event.returnValue = false; } " onpaste="return false;" ondragenter="return false;">
              <%else%>
                admin
              <%end if%>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��ʵ����<font color="#FF0000">*</font>&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=12 then%>
              <input name=truename size=15 maxlength=10 class="smallInput" value='<%=request("truename")%>'>
            <%else
              if rs("username")<>"admin" then%>
                <input name=truename size=15 maxlength=10 class="smallInput" value='<%=rs("name")%>'>
              <%else%>
                admin
              <%end if%>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" class="smallInput" align=right nowrap>����<font color="#FF0000">*</font>&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=12 then%>
              <input name=password size=15 type="password" maxlength=10 class="smallInput" value='<%=request("password")%>'>
            <%else%>
              <input name=password size=15 type="password" maxlength=10 class="smallInput" value='<%=rs("password")%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" class="smallInput" align=right nowrap>���ڵ�λ<font color="#FF0000">*</font>&nbsp;</td>
          <td align=left colspan=2>
	    ����
            <%if mode=12 then%> 
              <select name="unit_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(202,1)">
            <%else%>
              <select name="unit_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(203,1)">
            <%end if%>
            <%while not rs1.EOF 
              if trim(unit_code)="" then unit_code=trim(rs1("unit_code"))%>
              <option value="<%=trim(rs1("unit_code"))%>"<%if left(unit_code,unit_len1)=left(rs1("unit_code"),unit_len1) then %> selected <% end if %>><%=trim(rs1("unit_name"))%></option>
              <%rs1.MoveNext 
            WEND%>
            </select>
            <br>�շ�վ
            <%if mode=12 then%> 
              <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(202,2)">
            <%else%>
              <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(203,2)">
            <%end if%>
            <option value="" <%if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then %> selected <% end if %>></option>
            <%rsMX.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn_system, 1, 1
            while not rsMX.EOF
              'if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then unit_code=trim(rsMX("unit_code"))%>
              <option value="<%=trim(rsMX("unit_code"))%>"<%if left(unit_code,unit_len2)=left(rsMX("unit_code"),unit_len2) then %> selected <% end if %>><%=trim(rsMX("unit_name"))%></option>
              <%rsMX.MoveNext 
            WEND
            rsMX.close%>
            </select>
          </td>
        </tr>
        <%if register=1 then%>
        <%else%>
          <tr>
            <td bgcolor="#eeeeee" align=right nowrap>Ȩ��<font color="#FF0000">*</font>&nbsp;</td>
            <td align=left valign=center colspan=2>
              <%if mode=13 then
                if rs("username")<>"admin" then%>
                  <input type=checkbox name=power value='1' <%isPowerCheck rs("power"),1%>>���ݵǼ�<br>
                  <input type=checkbox name=power value='2' <%isPowerCheck rs("power"),2%>>���ݱ��<br>
                  <input type=checkbox name=power value='3' <%isPowerCheck rs("power"),3%>>��ѯͳ��<br>
                <%end if%>
              <%else%>
                <input type=checkbox name=power value='0'>ϵͳ����<br>
                <input type=checkbox name=power value='1'>���ݵǼ�<br>
                <input type=checkbox name=power value='2'>���ݱ��<br>
                <input type=checkbox name=power value='3'>��ѯͳ��<br>
              <%end if%>
            </td>
          </tr>
        <%end if%>
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
  else
    closedb()
  end if
  showctail
end sub

sub ShowInputForm3(ErrMsg)
  'on error resume next
  showchead()%>

  <form method="POST" action="muser.asp?mode=5&username=<%=username%>" name="input3">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        <%if username="" then%>
          [<a href="muser.asp?mode=1">����</a>]
        <%else%>
          [<a href="muser.asp?mode=8&username=<%=username%>">����</a>]
        <%end if%>
      </td>
    </tr>
    <tr><td><hr size="1" noshade width="100%"></td></tr>
    <tr><td>
      <table width="450" border="0" cellspacing="1" cellpadding="0" align="center" bordercolor="#FFCC33">
        <tr>
        <%if Trim(ErrMsg) <> "" then%>
          <td align=center><%=errmsg%><br><br></td>
        <%else%>
          <td align=center>�������������<br><br></td>
        <%end if%>
        </tr>
        <tr>
          <td align=center><input type="text" name="memname" size="60" maxlength="20" class="smallInput" value="<%=request("memname")%>"></td>
        </tr>
        <tr align="center">
          <td><br><input class="buttonface" type="submit" value=" ��ʼ���� " id=submit1 name=submit1></td>
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
  '��ʾ
  if not isEmpty(request("page")) then
    cpage = clng(request("page"))
  else
    cpage = 1
  end if
  showchead()
  'Response.Write "<br>"
  opendb()
   
  set rs=server.createobject("adodb.recordset")
  set rs1=server.createobject("adodb.recordset")
  set rsMX=server.createobject("adodb.recordset")
  rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn_system, 1, 1
  %>
  <form name="input1">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="#eeeeee" align=left>
      ����
      <select name="unit_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(301,1)">
      <%while not rs1.EOF 
        if trim(unit_code)="" then unit_code=trim(rs1("unit_code"))%>
        <option value="<%=trim(rs1("unit_code"))%>"<%if left(unit_code,unit_len1)=left(rs1("unit_code"),unit_len1) then %> selected <% end if %>><%=trim(rs1("unit_name"))%></option>
        <%rs1.MoveNext 
      WEND%>
      </select>
      �շ�վ
      <select name="unit_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(301,2)">
      <option value="" <%if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then %> selected <% end if %>></option>
      <%rsMX.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn_system, 1, 1
      while not rsMX.EOF
        'if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then unit_code=trim(rsMX("unit_code"))%>
        <option value="<%=trim(rsMX("unit_code"))%>"<%if left(unit_code,unit_len2)=left(rsMX("unit_code"),unit_len2) then %> selected <% end if %>><%=trim(rsMX("unit_name"))%></option>
        <%rsMX.MoveNext 
      WEND
      rsMX.close%>
      </select>
    </td>
    <%rs.open "select * from userinfo where unit_code='"+unit_code+"'",conn_system, 1, 1
    if rs.recordcount <> 0 then
      rs.movefirst
      rs.CacheSize = 5
      rs.PageSize = 10
      if cpage>rs.pagecount then cpage=1
      rs.AbsolutePage = cpage%>
      <td valign="bottom" align="right">��<%=cstr(cpage)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>λ�û�</td>
      <td align="right">
        [<a href="muser.asp?mode=2">���</a>]
        <%if cpage <> 1 then%>
          [<a href="muser.asp?mode=1&unit_code=<%=unit_code%>&page=<%=cstr(cpage-1)%>">��һҳ</a>]
        <%end if%>
        <%if cpage <> rs.PageCount then%>
          [<a href="muser.asp?mode=1&unit_code=<%=unit_code%>&page=<%=cstr(cpage+1)%>">��һҳ</a>]
        <%end if%>
        <%if rs.PageCount > 1 then%>
	  <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage then%>
                <option selected value="muser.asp?mode=1&unit_code=<%=unit_code%>&page=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="muser.asp?mode=1&unit_code=<%=unit_code%>&page=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
            <%next%>
          </select>
        <%end if%>
      </td>
     </tr>
     <tr>
       <td colspan="6">
         <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
           <tr bgcolor=<%=skincolor()%>>
	     <td width=60 align=center>�û���</td>
             <td width=140 align=center>���ڵ�λ</td>
             <td width=50 align=center>��ʵ����</td>
             <td align=center>Ȩ��</td>
             <td width=30 align=center>��Ч</td>
           </tr>
           <%fl = False
           for i = 1 to rs.PageSize
           if not rs.EOF then
             if fl then%>
               <tr bgcolor="#eeeeee">
             <%else%>
               <tr>
             <%end if%>
             <td><a href="muser.asp?mode=8&username=<%=rs("username")%>"><%=rs("username")%></a></td>
             <td><%response.write(rs("unit_name"))%></td>
             <td><%response.write(rs("name"))%></td>
             <td>
                <%if trim(rs("username"))="admin" then
                  s="ϵͳ����"
                else
                  s=""
                  if instr(rs("power"),",0,")<>0 then s=s+"ϵͳ����,"
                  if instr(rs("power"),",1,")<>0 then s=s+"���ݵǼ�,"
                  if instr(rs("power"),",2,")<>0 then s=s+"���ݱ��,"
                  if instr(rs("power"),",3,")<>0 then s=s+"��ѯͳ��,"
                  if s="" then s="δ֪,"
                end if
                Response.Write s%>
             </td>
             <td align=center>
               <%
               if trim(rs("username"))<>"admin" then
                 if rs("valid") then
                   response.write "��"
                 else
                   response.write "��"
                 end if
               else
                 Response.Write "��"
               end if%>
             </td>
             </tr>
             <%rs.MoveNext
             fl = not fl
           end if
           next%>
          </table>
        </td></tr>
      </table>
     </form>
  <%else%>
  <!--<br><br>-->
    <!--<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>-->
        <td valign="bottom" align="right"></td>     
        <td align="right">
          [<a href="muser.asp?mode=2">���</a>]
        </td>
      </tr>
      <tr><td colspan="6"><hr size=1 width=100% noshade></td></tr>
      <tr><td align="center" colspan="6"><font size="6">û�м�¼</font></td></tr>
    </table>
    </form>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=2 or mode=3 then
  '��Ӽ��޸�
  if request("memname")<>"" or request("password")<>"" then
    FoundError=false
    ErrMsg=""
    memname = trim(request("memname"))
    truename=trim(request("truename"))
    if request("unit_code2")="" then
      unit_code=request("unit_code1")
    else
      unit_code=request("unit_code2")
    end if
    FRMunit=trim(request("FRMunit"))
    if mode=2 then
      if memname = "" then
        ErrMsg="�������û���"
        foundError=True
      end if
      if truename = "" then
        ErrMsg="��������ʵ����"
        foundError=True
      end if
      fl = false
      for i=1 to len(username)
        s = mid(username,i,1)
        if Len(Hex(asc(s)))<=2 then
          if not ( (s>="0" and s=<"9") ) then fl = True
        end if
      next
      if fl then
        if ErrMsg <> "" then
          ErrMsg = ErrMsg + "<br>"
        else
          ErrMsg = "�û���������Ч�ַ�"
          foundError=True
        end if
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select username from userinfo where username='" + memname + "'", conn_system, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "�û����ظ�"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    else
      if memname="" then memname="admin"
      if memname = "" then
        ErrMsg="�������û���"
        foundError=True
      end if
      if truename = "" then
        ErrMsg="��������ʵ����"
        foundError=True
      end if
      fl = false
      for i=1 to len(username)
        s = mid(username,i,1)
        if Len(Hex(asc(s)))<=2 then
          if not ( (s>="0" and s=<"9") ) then fl = True
        end if
      next
      if fl then
        if ErrMsg <> "" then
          ErrMsg = ErrMsg + "<br>"
        else
          ErrMsg = "�û���������Ч�ַ�"
          foundError=True
        end if
      end if
      '���Ĺ����û����Ƿ����
      if username<>memname then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select username from userinfo where username='" + memname + "'", conn_system, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "�û����ظ�"
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
      s=""
      if register=1 then
      else
        for i=1 to request.form("power").count
          s=s+","+request.form("power")(i)+","
        next
      end if
      if mode = 2 then
        '�����
        opendb()
        set rs=server.createobject("adodb.recordset")
        unit_name=""
        rs.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"'",conn_system,1,1
        if rs.recordcount>0 then
          unit_name=unit_name+"["+rs("unit_name")+"]"
          rs.close
          rs.open "select * from unit where left(unit_code," & unit_len2 & ")='" & left(unit_code,unit_len2) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00'",conn_system,1,1
          if rs.recordcount>0 then
            unit_name=unit_name+"["+rs("unit_name")+"]"
            rs.close
            rs.open "select * from unit where unit_code='"+unit_code+"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00'",conn_system,1,1
            if rs.recordcount>0 then
              unit_name=unit_name+"["+rs("unit_name")+"]"
            end if
          end if
        end if
        rs.close
        rs.open "userinfo", conn_system, 1, 3
        rs.addnew
        rs("name")=truename
        rs("username")=memname
        rs("password")=request("password")
        rs("unit_code")=unit_code
        rs("unit_name")=unit_name
	rs("power")=s
        if register=1 then'ע���û�
          rs("valid")=0
        else
          rs("valid")=1
        end if
        rs("skin")="green"
        rs.update
        rs.close
        set rs=nothing
        closedb()
        'response.write "muser.asp?mode=8&username=" & memname & "&register=" & register
        Response.Redirect "muser.asp?mode=8&username=" & memname&"&register=" & register
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        unit_name=""
        rs.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"'",conn_system,1,1
        if rs.recordcount>0 then
          unit_name=unit_name+"["+rs("unit_name")+"]"
          rs.close
          rs.open "select * from unit where left(unit_code," & unit_len2 & ")='" & left(unit_code,unit_len2) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00'",conn_system,1,1
          if rs.recordcount>0 then
            unit_name=unit_name+"["+rs("unit_name")+"]"
            rs.close
            rs.open "select * from unit where unit_code='"+unit_code+"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00'",conn_system,1,1
            if rs.recordcount>0 then
              unit_name=unit_name+"["+rs("unit_name")+"]"
            end if
          end if
        end if
        rs.close
        sql="update userinfo set username='"+memname+"',name='"+truename+"',password='"+request("password")+"',unit_code='"+unit_code+"',unit_name='"+unit_name+"',power='"+s+"'  where username='"+username+"'"
        Response.Write sql
        conn_system.Execute sql
        set rs=nothing
        closedb()
        Response.Redirect "muser.asp?mode=8&username=" & username & "&register=" & register
      end if
    end if
  else
      ShowInputForm1 mode,""
  end if

elseif mode=4 then
  'ɾ��ȷ��
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        <%if username="" then%>
          [<a href="muser.asp?mode=1">����</a>]
        <%else%>
          [<a href="muser.asp?mode=8&username=<%=username%>">����</a>]
        <%end if%>
     </td>
    </tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      <%if isempty(request("username")) then%>
       �Բ��𣬴���Ĳ��������������ء�
      <%else%>
      ���Ҫɾ���û���<%=username%>����
      <br><br>
      [<a href="muser.asp?mode=7&username=<%=username%>&nusername=<%=request("username")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="muser.asp?mode=8&username=<%=username%>">����</a>]
      <%end if%>
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()
elseif mode=5 then
  '����
  if trim(request("memname")) <> "" then
    opendb()
    set rs=server.createobject("adodb.recordset")
    sql=""
    if trim(request("memname")) <> "" then
      sql="(username like '%" + trim(request("memname")) + "%')"
    end if
    rs.open "select * from userinfo where " + sql, conn_system, 1, 1
    if rs.recordcount=0 then
      rs.close
      set rs=nothing
      closedb()
      showinputform3 "Can't find any match record, please reinput search condition."
    else
      showchead()%>
      <br>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td align="right">
            <%if username="" then%>
              [<a href="muser.asp?mode=1">����</a>] 
            <%else%>
              [<a href="muser.asp?mode=8&username=<%=username%>">����</a>] 
            <%end if%>
            [<a href="muser.asp?mode=5&username=<%=username%>">��������</a>]
         </td>
        </tr>
        <tr><td>
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=60 align=center>�û���</td>
              <td width=140 align=center>���ڵ�λ</td>
              <td width=50 align=center>��ʵ����</td>
              <td align=center>Ȩ��</td>
              <td width=30 align=center>��Ч</td>
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
                <td><a href="muser.asp?mode=8&username=<%=rs("username")%>"><%=rs("username")%></a></td>
                <td><%response.write(rs("unit_name"))%></td>
                <td><%response.write(rs("name"))%></td>
                <td><%
                if trim(rs("username"))="admin" then
                  s="ϵͳ����"
                else
                  s=""
                  if instr(rs("power"),",0,")<>0 then s=s+"ϵͳ����,"
                  if instr(rs("power"),",1,")<>0 then s=s+"���ݵǼ�,"
                  if instr(rs("power"),",2,")<>0 then s=s+"���ݱ��,"
                  if instr(rs("power"),",3,")<>0 then s=s+"��ѯͳ��,"
                  if s="" then s="δ֪,"
                end if
                Response.Write s
                %></td>
                <td align=center>
                <%
                if trim(rs("username"))<>"admin" then
                  if rs("valid") then
                    response.write "��"
                  else
                    response.write "��"
                  end if
                else
                  Response.Write "��"
                end if
                %>
                </td>
              </tr>
              <%rs.MoveNext
              fl = not fl
              end if
            next%>
          </table>
        </td></tr>
      </table>
      <%
      rs.close
      set rs=nothing
      closedb()
      showctail()
      if not isempty(request("memname")) then
        session("cond1") = trim(request("memaname"))
      else
        session("cond1") = ""
      end if
      if not isempty(request("email")) then
        session("cond2") = trim(request("email"))
      else
        session("cond2") = ""
      end if
    end if
  else
    ShowInputForm3 ""
  end if

elseif mode=6 then
  'change state
  opendb()
  set rs=server.createobject("adodb.recordset")
  rs.open "select valid from userinfo where username='"+request("username")+"'", conn_system, 1, 3
  if rs.recordcount<>0 then
    if rs("valid") then
      conn_system.Execute "update userinfo set valid=0 where username='"+request("username")+"'"
    else
      conn_system.Execute "update userinfo set valid=1 where username='"+request("username")+"'"
    end if
    rs.update
  end if
  rs.close
  closedb()
  response.redirect "muser.asp?mode=8&username="+request("username")

elseif mode=7 then
  'delete
  opendb()
  conn_system.execute "delete from userinfo where username='" + request("username")+"'"
  closedb()
  delaySecond(2)
  Response.Redirect ("muser.asp?mode=8&username=" + request("username"))
elseif mode=8 then
  '��ʾ
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  rs.open "select * from userinfo order by username", conn_system, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    if username = "" then
      username = rs("username")
    else
      rs.Find "username = '" + username +"'"
      if rs.EOF then
        rs.movefirst
        username = rs("username")
      end if
    end if
    %>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <%if register=1 then%>
        <td algin="center">��ϲ����ע��ɹ���</td>
        <td align="right">[<a href="muser.asp?mode=3&register=<%=register%>&username=<%=rs("username")%>&unit_code=<%=rs("unit_code")%>">�༭</a>]</td>
        <%else%>
        <td>
          ��<strong><font color="blue"><%=rs.recordcount%></font></strong>λ�û� [<a href='muser.asp?mode=1'>�û��б�</a>]
          <%
          rs.MovePrevious
          if not rs.BOF then
            Response.Write "[<a href='muser.asp?mode=8&username="+rs("username")+"'>��һҳ</a>] "
          end if
          rs.Move 2
          if not rs.EOF then
            Response.Write "[<a href='muser.asp?mode=8&username="+rs("username")+"'>��һҳ</a>]"
          end if
          rs.MovePrevious
          %>
        </td>
        <td align="right">
          [<a href="muser.asp?mode=2&username=<%=username%>">���</a>]
          <%if rs("username")<>"admin" then%>
            [<a href="muser.asp?mode=3&username=<%=rs("username")%>&unit_code=<%=rs("unit_code")%>">�༭</a>]
            <%rs.MoveNext
            if not rs.EOF then
              Response.Write "[<a href='muser.asp?mode=4&username="+username+"&nusername="+rs("username")+"'>ɾ��</a>] "
              rs.MovePrevious
            else
              rs.Move -2
              if not rs.BOF then
                Response.Write "[<a href='muser.asp?mode=4&username="+username+"&nusername="+rs("username")+"'>ɾ��</a>] "
              else
                Response.Write "[<a href='muser.asp?mode=4&username="+username+"&nusername=n/a'>ɾ��</a>] "
              end if
              rs.MoveNext
            end if
          end if
          %>
          [<a href="muser.asp?mode=5&username=<%=username%>">����</a>]
        </td>
        <%end if%>
      </tr>
      <tr><td colspan="2"><hr noshade size="1" width="100%"></td></tr>
      <tr><td colspan="2">
        <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
            <%if register=1 then%>
            <%else%>
            <tr>
              <td bgcolor="#eeeeee" align=right width=20%>��Ч&nbsp;</td>
              <td align=left colspan=2><font color=red>
              <%
              if rs("username")<>"admin" then
                if rs("valid") then
                  response.write "&nbsp;��"
                else
                  response.write "&nbsp;��"
                end if%>
                </font>&nbsp;[<a href="muser.asp?mode=6&username=<%=username%>">�ı�</a>]
              <%else
              Response.Write "YES"
              end if
              %>
              </td>
            </tr>
            <%end if%>
            <tr>
              <td bgcolor="#eeeeee" align=right nowrap>�û���&nbsp;</td>
              <td align=left colspan=2>&nbsp;<%=rs("username")%></td>
            </tr>
            <tr>
              <td bgcolor="#eeeeee" align=right nowrap>���ڵ�λ&nbsp;</td>
              <td align=left colspan=2>&nbsp;<%response.write(rs("unit_name"))%></td>
            </tr>
            <tr>
              <td bgcolor="#eeeeee" align=right nowrap>��ʵ����&nbsp;</td>
              <td align=left colspan=2>&nbsp;<%response.write(rs("name"))%></td>
            </tr>
            <%if register=1 then%>
            <%else%>
            <tr>
              <td bgcolor="#eeeeee" align=right nowrap>Ȩ��&nbsp;</td>
              <td align=left colspan=2>
                <%drawPowerCheck rs("power"),0%>ϵͳ����<br>
                <%drawPowerCheck rs("power"),1%>���ݵǼ�<br>
                <%drawPowerCheck rs("power"),2%>���ݱ��<br>
                <%drawPowerCheck rs("power"),3%>��ѯͳ��<br>
              </td>
            </tr>
            <%end if%>
        </table>
      </td></tr>
      <tr><td colspan="2"><hr noshade size="1" width="100%"></td></tr>
    </table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <%if register=1 then%>
      <%else%>
      <tr>
        <td align="right">
          [<a href="muser.asp?mode=2">���</a>]
        </td>
      </tr>
      <%end if%>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">û�м�¼</font></td></tr>
      <tr><td>&nbsp;</td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()
elseif mode=12 or mode=13 then
  '��Ӽ��޸�(δ���)
  if request("memname")<>"" or request("password")<>"" then
    FoundError=false
    ErrMsg=""
    memname = trim(request("memname"))
    truename=trim(request("truename"))
    if request("unit_code2")="" then
      unit_code=request("unit_code1")
    else
      unit_code=request("unit_code2")
    end if
    FRMunit=trim(request("FRMunit"))
    if mode=12 then
      if memname = "" then
        ErrMsg="�������û���"
        foundError=True
      end if
      if truename = "" then
        ErrMsg="��������ʵ����"
        foundError=True
      end if
      fl = false
      for i=1 to len(username)
        s = mid(username,i,1)
        if Len(Hex(asc(s)))<=2 then
          if not ( (s>="0" and s=<"9") ) then fl = True
        end if
      next
      if fl then
        if ErrMsg <> "" then
          ErrMsg = ErrMsg + "<br>"
        else
          ErrMsg = "�û���������Ч�ַ�"
          foundError=True
        end if
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select username from userinfo where username='" + memname + "'", conn_system, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "�û����ظ�"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    else
      if memname="" then memname="admin"
      if memname = "" then
        ErrMsg="�������û���"
        foundError=True
      end if
      if truename = "" then
        ErrMsg="��������ʵ����"
        foundError=True
      end if
      fl = false
      for i=1 to len(username)
        s = mid(username,i,1)
        if Len(Hex(asc(s)))<=2 then
          if not ( (s>="0" and s=<"9") ) then fl = True
        end if
      next
      if fl then
        if ErrMsg <> "" then
          ErrMsg = ErrMsg + "<br>"
        else
          ErrMsg = "�û���������Ч�ַ�"
          foundError=True
        end if
      end if
      '���Ĺ����û����Ƿ����
      if username<>memname then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select username from userinfo where username='" + memname + "'", conn_system, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "�û����ظ�"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm11 mode,errmsg
    else
      s=""
      if register=1 then
      else    
        for i=1 to request.form("power").count
          s=s+","+request.form("power")(i)+","
        next
      end if
      if mode = 12 then
        '�����
        opendb()
        set rs=server.createobject("adodb.recordset")
        unit_name=""
        rs.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"'",conn_system,1,1
        if rs.recordcount>0 then
          unit_name=unit_name+"["+rs("unit_name")+"]"
          rs.close
          rs.open "select * from unit where left(unit_code," & unit_len2 & ")='" & left(unit_code,unit_len2) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00'",conn_system,1,1
          if rs.recordcount>0 then
            unit_name=unit_name+"["+rs("unit_name")+"]"
            rs.close
            rs.open "select * from unit where unit_code='"+unit_code+"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00'",conn_system,1,1
            if rs.recordcount>0 then
              unit_name=unit_name+"["+rs("unit_name")+"]"
            end if
          end if
        end if
        rs.close
        rs.open "userinfo", conn_system, 1, 3
        rs.addnew
        rs("name")=truename
        rs("username")=memname
        rs("password")=request("password")
        rs("unit_code")=unit_code
        rs("unit_name")=unit_name
	rs("power")=s
        if register=1 then'ע���û�
          rs("valid")=0
        else
          rs("valid")=1
        end if
        rs("skin")="green"
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "muser.asp?mode=18&username=" & memname&"&register=" & register
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        unit_name=""
        rs.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"'",conn_system,1,1
        if rs.recordcount>0 then
          unit_name=unit_name+"["+rs("unit_name")+"]"
          rs.close
          rs.open "select * from unit where left(unit_code," & unit_len2 & ")='" & left(unit_code,unit_len2) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00'",conn_system,1,1
          if rs.recordcount>0 then
            unit_name=unit_name+"["+rs("unit_name")+"]"
            rs.close
            rs.open "select * from unit where unit_code='"+unit_code+"' and mid(unit_code,"& (unit_len2+1) & "," & (unit_len3-unit_len2) & ")<>'00'",conn_system,1,1
            if rs.recordcount>0 then
              unit_name=unit_name+"["+rs("unit_name")+"]"
            end if
          end if
        end if
        rs.close
        sql="update userinfo set username='"+memname+"',name='"+truename+"',password='"+request("password")+"',unit_code='"+unit_code+"',unit_name='"+unit_name+"',power='"+s+"'  where username='"+username+"'"
        Response.Write sql
        conn_system.Execute sql
        set rs=nothing
        closedb()
        Response.Redirect "muser.asp?mode=18&username=" & username & "&register=" & register
      end if
    end if
  else
      ShowInputForm11 mode,""
  end if
elseif mode=14 then
  'ɾ��ȷ��
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        <%if username="" then%>
          [<a href="muser.asp?mode=1">����</a>]
        <%else%>
          [<a href="muser.asp?mode=18&username=<%=username%>">����</a>]
        <%end if%>
     </td>
    </tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      <%if isempty(request("username")) then%>
       �Բ��𣬴���Ĳ��������������ء�
      <%else%>
      ���Ҫɾ���û���<%=username%>����
      <br><br>
      [<a href="muser.asp?mode=17&username=<%=username%>&nusername=<%=request("username")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="muser.asp?mode=18&username=<%=username%>">����</a>]
      <%end if%>
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%showctail()
elseif mode=15 then
  '����
  if trim(request("memname")) <> "" then
    opendb()
    set rs=server.createobject("adodb.recordset")
    sql=""
    if trim(request("memname")) <> "" then
      sql="(username like '%" + trim(request("memname")) + "%')"
    end if
    rs.open "select * from userinfo where " + sql, conn_system, 1, 1
    if rs.recordcount=0 then
      rs.close
      set rs=nothing
      closedb()
      showinputform3 "Can't find any match record, please reinput search condition."
    else
      showchead()%>
      <br>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td align="right">
            <%if username="" then%>
              [<a href="muser.asp?mode=1">����</a>] 
            <%else%>
              [<a href="muser.asp?mode=18&username=<%=username%>">����</a>] 
            <%end if%>
            [<a href="muser.asp?mode=15&username=<%=username%>">��������</a>]
         </td>
        </tr>
        <tr><td>
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=60 align=center>�û���</td>
              <td width=140 align=center>���ڵ�λ</td>
              <td width=50 align=center>��ʵ����</td>
              <td align=center>Ȩ��</td>
              <td width=30 align=center>��Ч</td>
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
                <td><a href="muser.asp?mode=8&username=<%=rs("username")%>"><%=rs("username")%></a></td>
                <td><%response.write(rs("unit_name"))%></td>
                <td><%response.write(rs("name"))%></td>
                <td><%
                if trim(rs("username"))="admin" then
                  s="ϵͳ����"
                else
                  s=""
                  if instr(rs("power"),",0,")<>0 then s=s+"ϵͳ����,"
                  if instr(rs("power"),",1,")<>0 then s=s+"���ݵǼ�,"
                  if instr(rs("power"),",2,")<>0 then s=s+"���ݱ��,"
                  if instr(rs("power"),",3,")<>0 then s=s+"��ѯͳ��,"
                  if s="" then s="δ֪,"
                end if
                Response.Write s
                %></td>
                <td align=center>
                <%
                if trim(rs("username"))<>"admin" then
                  if rs("valid") then
                    response.write "��"
                  else
                    response.write "��"
                  end if
                else
                  Response.Write "��"
                end if
                %>
                </td>
              </tr>
              <%rs.MoveNext
              fl = not fl
              end if
            next%>
          </table>
        </td></tr>
      </table>
      <%
      rs.close
      set rs=nothing
      closedb()
      showctail()
      if not isempty(request("memname")) then
        session("cond1") = trim(request("memaname"))
      else
        session("cond1") = ""
      end if
      if not isempty(request("email")) then
        session("cond2") = trim(request("email"))
      else
        session("cond2") = ""
      end if
    end if
  else
    ShowInputForm3 ""
  end if

elseif mode=16 then
  'change state
  opendb()
  set rs=server.createobject("adodb.recordset")
  rs.open "select valid from userinfo where username='"+request("username")+"'", conn_system, 1, 3
  if rs.recordcount<>0 then
    if rs("valid") then
      conn_system.Execute "update userinfo set valid=0 where username='"+request("username")+"'"
    else
      conn_system.Execute "update userinfo set valid=1 where username='"+request("username")+"'"
    end if
    rs.update
  end if
  rs.close
  closedb()
  response.redirect "muser.asp?mode=18&username="+request("username")

elseif mode=17 then
  'delete
  opendb()
  conn_system.execute "delete from userinfo where username='" + request("username")+"'"
  closedb()
  delaySecond(2)
  Response.Redirect ("muser.asp?mode=18&username=" + request("username"))
elseif mode=18 then
  '��ʾδ��˵��û�
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  rs.open "select * from userinfo where valid<>1 order by username", conn_system, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    if username = "" then
      username = rs("username")
    else
      rs.Find "username = '" + username +"'"
      if rs.EOF then
        rs.movefirst
        username = rs("username")
      end if
    end if
    %>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td>
          ��<strong><font color="blue"><%=rs.recordcount%></font></strong>λδ��˵��û�
          <%
          rs.MovePrevious
          if not rs.BOF then
            Response.Write "[<a href='muser.asp?mode=18&username="+rs("username")+"'>��һҳ</a>] "
          end if
          rs.Move 2
          if not rs.EOF then
            Response.Write "[<a href='muser.asp?mode=18&username="+rs("username")+"'>��һҳ</a>]"
          end if
          rs.MovePrevious
          %>
        </td>
        <td align="right">
          <%if rs("username")<>"admin" then%>
            [<a href="muser.asp?mode=13&username=<%=rs("username")%>&unit_code=<%=rs("unit_code")%>">�༭</a>]
            <%rs.MoveNext
            if not rs.EOF then
              Response.Write "[<a href='muser.asp?mode=14&username="+username+"&nusername="+rs("username")+"'>ɾ��</a>] "
              rs.MovePrevious
            else
              rs.Move -2
              if not rs.BOF then
                Response.Write "[<a href='muser.asp?mode=14&username="+username+"&nusername="+rs("username")+"'>ɾ��</a>] "
              else
                Response.Write "[<a href='muser.asp?mode=14&username="+username+"&nusername=n/a'>ɾ��</a>] "
              end if
              rs.MoveNext
            end if
          end if
          %>
          [<a href="muser.asp?mode=15&username=<%=username%>">����</a>]
        </td>
      </tr>
      <tr><td colspan="2"><hr noshade size="1" width="100%"></td></tr>
      <tr>
        <td colspan="2">
          <table width="500" border="0" cellspacing="1" cellpadding="1" align="center">
            <tr>
              <td bgcolor="#eeeeee" align=right width=20%>��Ч&nbsp;</td>
              <td align=left colspan=2><font color=red>
              <%
              if rs("username")<>"admin" then
                if rs("valid") then
                  response.write "&nbsp;��"
                else
                  response.write "&nbsp;��"
                end if%>
                </font>&nbsp;[<a href="muser.asp?mode=16&username=<%=username%>">�ı�</a>]
              <%else
              Response.Write "YES"
              end if
              %>
              </td>
            </tr>
            <tr>
              <td bgcolor="#eeeeee" align=right nowrap>�û���&nbsp;</td>
              <td align=left colspan=2>&nbsp;<%=rs("username")%></td>
            </tr>
            <tr>
              <td bgcolor="#eeeeee" align=right nowrap>���ڵ�λ&nbsp;</td>
              <td align=left colspan=2>&nbsp;<%response.write(rs("unit_name"))%></td>
            </tr>
            <tr>
              <td bgcolor="#eeeeee" align=right nowrap>��ʵ����&nbsp;</td>
              <td align=left colspan=2>&nbsp;<%response.write(rs("name"))%></td>
            </tr>
            <tr>
              <td bgcolor="#eeeeee" align=right nowrap>Ȩ��&nbsp;</td>
              <td align=left colspan=2>
                <%drawPowerCheck rs("power"),0%>ϵͳ����<br>
                <%drawPowerCheck rs("power"),1%>���ݵǼ�<br>
                <%drawPowerCheck rs("power"),2%>���ݱ��<br>
                <%drawPowerCheck rs("power"),3%>��ѯͳ��<br>
              </td>
            </tr>
        </table>
      </td></tr>
      <tr><td colspan="2"><hr noshade size="1" width="100%"></td></tr>
    </table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          [<a href="muser.asp?mode=1">����</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">û��δ��˼�¼</font></td></tr>
      <tr><td>&nbsp;</td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()
end if

function GetUnitName(s)
  dim rsMX
  sql="select * from unit where unit_code='"&left(s,6)&"'"
  set rsMX=conn_system.execute(sql)
  if rsMX.eof and rsMX.bof then
    rsMX.close:set rsMXs=nothing
    getunitname="":exit function
  else
    getunitname=rsMX("unit_name")
  end if
  rsMX.close
  if len(s)>6 then
    sql="select * from station where station_code='"&left(s,8)&"'"
    set rsMX=conn_system.execute(sql)
    if rsMX.eof and rsMX.bof then
      rsMX.close:set rsMX=nothing
      exit function
    else
      getunitname=getunitname+rsMX("station_name")
    end if
    rsMX.close
    if len(s)>8 then
      sql="select * from section where section_code='"&left(s,10)&"'"
      set rsMX=conn_system.execute(sql)
      if rsMX.eof and rsMX.bof then
        rsMX.close:set rsMX=nothing
   	exit function
      else
	getunitname=getunitname+rsMX("section_name")
      end if
      rsMX.close
    end if
  end if
end function
%>