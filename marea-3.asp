<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if session("username")="" or instr(session("power"),",0,")=0 then
    Response.Redirect "notlogin.asp"
end if
%>

<!--#include file="fcommon.asp"-->

<%
dim conn, mode, username, rs, sql, errmsg, founderror, s, t, i, fl, dq,odq, cpage,kpbm,st,dwxh

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
  if not isEmpty(request("page")) and isnumeric(request("page")) then
    cpage = clng(request("page"))
  else
    cpage = 1
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
  <title>文章类别管理</title>
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
      <b>文章类别设置</b>
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
    <form method="POST" action="marea-3.asp?mode=2&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from wzlb where wzlb_name='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-3.asp?mode=3&page=<%=cpage%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-3.asp?mode=1&page=<%=cpage%>">返回</a>]
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
            <td colspan="3">请输入文章类别，然后点击“OK”</td>
          <%else%>
            <td colspan="3">请编辑文章类别，然后点击“OK”</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>文章类别&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=dq size=25 maxlength=20 class="smallInput" value='<%=request("dq")%>'>
            <%else%>
              <input name=dq size=25 maxlength=20 class="smallInput" value='<%=trim(rs("wzlb_name"))%>'>
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

sub ShowInputForm3(ErrMsg)
  'on error resume next
  showchead()%>

  <form method="POST" action="marea-3.asp?mode=5&username=<%=username%>" name="input3">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-3.asp?mode=8&username=<%=username%>">返回</a>]
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
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  rs.open "select * from wzlb order by wzlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage>rs.pagecount then cpage=1
    rs.AbsolutePage = cpage
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">第<%=cstr(cpage)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个文章类别</td>
          <td align="right">
          [<a href="marea-3.asp?mode=2">添加</a>]
          <%if cpage <> 1 then%>
            [<a href="marea-3.asp?mode=1&page=<%=cstr(cpage-1)%>">上一页</a>]
          <%end if%>
          <%if cpage <> rs.PageCount then%>
            [<a href="marea-3.asp?mode=1&page=<%=cstr(cpage+1)%>">下一页</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage then%>
                <option selected value="marea-3.asp?mode=1&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="marea-3.asp?mode=1&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=50% align=center>文章类别</td>
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
                <td align=center><%=trim(rs("wzlb_name"))%></td>
                <td align=center>
                  <a href="marea-3.asp?mode=3&page=<%=cpage%>&odq=<%=trim(rs("wzlb_name"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-3.asp?mode=4&page=<%=cpage%>&dq=<%=trim(rs("wzlb_name"))%>&dwxh=<%=trim(rs("wzlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("wzlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-3.asp?mode=8&page=<%=cpage%>&dq=<%=trim(rs("wzlb_code"))%>&sort=up&dwxh=<%=trim(rs("wzlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("wzlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-3.asp?mode=8&page=<%=cpage%>&dq=<%=trim(rs("wzlb_code"))%>&sort=down&dwxh=<%=trim(rs("wzlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
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
          [<a href="marea-3.asp?mode=2">添加</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()

elseif mode=2 or mode=3 then
  '添加及修改
  if request("dq")<>"" then
    FoundError=false
    ErrMsg=""
    dq = trim(request("dq"))
    if mode=2 then
      if dq = "" then
        ErrMsg="请输入文章类别"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册，判断有无重复的同一时间做两件事
        rs.open "select wzlb_name from wzlb where wzlb_name='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "文章类别重复"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select wzlb_name from wzlb", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
    else
      '看改过的用户名是否存在
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select wzlb_name from wzlb where wzlb_name='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "文章类别重复"
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
        rs.open "select wzlb_code from wzlb order by wzlb_code desc", conn,1,1'生成编号
	if rs.recordcount=0 then
	  kpbm="01"
	else
	  rs.movefirst
	  st=cstr(cint(rs("wzlb_code"))+1)
	  for i=len(st) to 1
	    st="0"&st
	  next
	  kpbm=kpbm&st
	end if
	rs.close
        rs.open "wzlb", conn, 1, 3
        rs.addnew
        rs("wzlb_code")=kpbm
        rs("wzlb_name")=dq
        rs("wzlb_sxh")=dwxh
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-3.asp?mode=1"
      else
        opendb()
        conn.Execute "update wzlb set wzlb_name='"+dq+"' where wzlb_name='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-3.asp?mode=1&page=" & cpage
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
        [<a href="marea-3.asp?mode=1">返回</a>]
     </td>
    </tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除文章类别“<%=request("dq")%>”？
      <br><br>
      [<a href="marea-3.asp?mode=7&page=<%=cpage%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-3.asp?mode=1&page=<%=cpage%>">算了</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()
elseif mode=5 then
  '搜索
  if trim(request("dq")) <> "" then
    opendb()
    set rs=server.createobject("adodb.recordset")
    sql=""
    if trim(request("dq")) <> "" then
      sql="(wzlb_name like '%" + trim(request("dq")) + "%')"
    end if
    rs.open "select * from wzlb where " + sql, conn, 1, 1
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
            [<a href="marea-3.asp?mode=1&username=<%=username%>">返回</a>] [<a href="marea-3.asp?mode=5&username=<%=username%>">继续查找</a>]
         </td>
        </tr>
        <tr><td>
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=50% align=center>文章类别</td>
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
                <td align=center><%=trim(rs("wzlb_name"))%></td>
                <td align=center>
                  <a href="marea-3.asp?mode=3&odq=<%=trim(rs("wzlb_name"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-3.asp?mode=4&dq=<%=trim(rs("wzlb_name"))%>"><img src="./images/del.gif" border=0></a>
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
      if not isempty(request("dq")) then
        session("cond1") = trim(request("dq"))
      else
        session("cond1") = ""
      end if
    end if
  else
    ShowInputForm3 ""
  end if

elseif mode=7 then
  'delete
  opendb()
  conn.execute "delete from wzlb where wzlb_name='" + request("dq")+"'"
  conn.execute "update wzlb set wzlb_sxh=wzlb_sxh-1 where wzlb_sxh>=" & request("dwxh")' 后面的顺序号往前推
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-3.asp?mode=1&page=" & cpage)

elseif mode=8 then
  'delete 上移/下移
  opendb()
  if request("sort")="up" then'上移
    conn.execute "update wzlb set wzlb_sxh=wzlb_sxh+1 where wzlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update wzlb set wzlb_sxh=wzlb_sxh-1 where wzlb_code='" + request("dq")+"'"
  else'下移
    conn.execute "update wzlb set wzlb_sxh=wzlb_sxh-1 where wzlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update wzlb set wzlb_sxh=wzlb_sxh+1 where wzlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-3.asp?mode=1&page=" & cpage)

end if
%>    