<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<!--#include file="./fcommon.asp"-->

<%
if session("username")=""  or instr(session("power"),",3,")=0 then
  Response.Redirect("notlogin.asp")
end if

dim conn, rs, rs1,rsMX, sql, errmsg, founderror, i, str1, mode, cpage, fl

if not isempty(request("mode")) and isnumeric(request("mode")) then
    mode = clng(request("mode"))
else
    mode=1
end if

sub opendb()
  set conn=server.createobject("ADODB.CONNECTION")
  if err.number<>0 then
    err.clear
    Response.Redirect "error.asp?errid=1"
  else
    conn.open sysconstr
    if err then
      err.clear
      Response.Redirect "error.asp?errid=1"
    end if
  end if
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
  <title>省交通厅、物价局</title>
  <link rel="stylesheet" type="text/css" href="/main.css">
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
  <body leftmargin="0" topmargin="0">
<%noRightClick()
end sub

sub showctail()
%>
  </body>
  </html>
<%
end sub

function initstr(str)
  dim i,s,t,fl
  s=trim(str)
  fl=false
  for i=1 to len(s)
    if mid(s,i,1)=" " then
      if not fl then
        t=t+" "
        fl=true
      end if
    else
      fl=false
      if mid(s,i,1)="+" then
        t=t+"%"
      else
        t=t+mid(s,i,1)
      end if
    end if
  next
  initstr=t
end function

sub ShowInputForm3(errmsg)
  'on error resume next
  showchead()%>
  <form method="POST" action="searchnyc.asp?mode=1" name="input3">
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>省交通厅、物价局查询</b></td>
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
          <td align=center><input type="text" name="dq" size="60" maxlength="20" class="smallInput" value=""></td>
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
  '搜索
  if trim(request("dq")) <> "" then
    if not isEmpty(request("page")) and isnumeric(request("page")) then
      cpage = clng(request("page"))
    else
      cpage = 1
    end if
    opendb()
    set rs=server.createobject("adodb.recordset")
    sql=""
    if trim(request("dq")) <> "" then
      sql="(car_model like '%" + trim(request("dq")) + "%')"
    end if
    rs.open "select * from dtxb_model where " + sql, conn, 1, 1
    if rs.recordcount=0 then
      rs.close
      set rs=nothing
      closedb()
      showinputform3 "Can't find any match record, please reinput search condition."
    else
      rs.movefirst
      rs.CacheSize = 5
      rs.PageSize = 10
      if cpage>rs.pagecount then cpage=1
      rs.AbsolutePage = cpage
      showchead()%>
      <br>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td align="right" colspan="2">
            [<a href="searchnyc.asp?mode=1">继续查找</a>]
          </td>
        </tr>
        <tr>
          <td valign="bottom">第<%=cstr(cpage)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个参数更正记录</td>
          <td align="right">
            <%if cpage <> 1 then%>
              [<a href="searchnyc.asp?mode=1&dq=<%=request("dq")%>&page=<%=cstr(cpage-1)%>">上一页</a>]
            <%end if%>
            <%if cpage <> rs.PageCount then%>
              [<a href="searchnyc.asp?mode=1&dq=<%=request("dq")%>&page=<%=cstr(cpage+1)%>">下一页</a>]
            <%end if%>
            <%if rs.PageCount > 1 then%>
              <select name="select2"  onchange="goto(this)">
                <%for i = 1 to rs.PageCount%>
                  <%if i = cpage then%>
                    <option selected value="searchnyc.asp?mode=1&dq=<%=request("dq")%>&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
                  <%else%>
                    <option value="searchnyc.asp?mode=1&dq=<%=request("dq")%>&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
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
                <td align=center>产品型号及名称</td>
				<td width=140 align=center>产地</td>
                <td width=100 align=center>原标定载质量</td>
                <td width=100 align=center>更正后载质量</td>
                <td width=100 align=center>原标定整备质量</td>
                <td width=100 align=center>更正后整备质量</td>
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
				  <td>
				    <%
					if isnull(rs("car_model")) then
					  response.write "&nbsp;"
					elseif trim(rs("car_model"))="" then
					  response.write "&nbsp;"
					else
					  response.write rs("car_model")
                    end if					
					%>
				  </td>
				  <td>
				    <%
					if isnull(rs("factoryname")) then
					  response.write "&nbsp;"
					elseif trim(rs("factoryname"))="" then
					  response.write "&nbsp;"
					else
					  response.write rs("factoryname")
                    end if					
					%>
				  </td>
				  <td>
				    <%
					if isnull(rs("o_lw")) then
					  response.write "&nbsp;"
					elseif trim(rs("o_lw"))="" then
					  response.write "&nbsp;"
					else
					  response.write rs("o_lw")
                    end if					
					%>
				  </td>
				  <td>
				    <%
					if isnull(rs("n_lw")) then
					  response.write "&nbsp;"
					elseif trim(rs("n_lw"))="" then
					  response.write "&nbsp;"
					else
					  response.write rs("n_lw")
                    end if					
					%>
				  </td>
				  <td>
				    <%
					if isnull(rs("o_ew")) then
					  response.write "&nbsp;"
					elseif trim(rs("o_ew"))="" then
					  response.write "&nbsp;"
					else
					  response.write rs("o_ew")
                    end if					
					%>
				  </td>
				  <td>
				    <%
					if isnull(rs("n_ew")) then
					  response.write "&nbsp;"
					elseif trim(rs("n_ew"))="" then
					  response.write "&nbsp;"
					else
					  response.write rs("n_ew")
                    end if					
					%>
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
      <%
      rs.close
      set rs=nothing
      closedb()
      showctail()
    end if
  else
    ShowInputForm3 ""
  end if
end if
%>    