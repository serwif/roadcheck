<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<!--#include file="./fcommon.asp"-->

<%
'if session("username")=""  or instr(session("power"),",3,")=0 then
'  Response.Redirect("notlogin.asp")
'end if

dim conn, rs, rs1,rsMX,rs2,rs3, sql,sql1,sql2,sql3,sql4,sql5, errmsg, founderror, i, str1, mode, cpage, fl,dwx,unit_code,shj1,shj2,byhj,bnlj,bbjzdwmc,bbjzzg,bbjzfh,bbjzzb,dq

if not isempty(request("mode")) and isnumeric(request("mode")) then
    mode = clng(request("mode"))
else
    mode=2
end if
if not isempty(request("unit_code")) then
    unit_code = request("unit_code")
else
    unit_code = ""
end if
if not isempty(request("dq")) then
    dq = request("dq")
else
    dq = ""
end if
if not isempty(request("shj1")) and isnumeric(request("shj1")) then
    shj1 = clng(request("shj1"))
else
    shj1=year(now)
end if
if not isempty(request("shj2")) and isnumeric(request("shj2")) then
    shj2 = clng(request("shj2"))
else
    shj2=0
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
  <title>花名册查询</title>
  <link rel="stylesheet" type="text/css" href="/main.css">
  </head>
  <script LANUGAGE="JavaScript">
  <!--
  function surfto(list){
   var myindex1=list.selectedIndex;
   if (myindex1!=0 & myindex1!=1){ location.href=list.options[list.selectedIndex].value }
  }
  function goto(list){
   //alert(list.options[list.selectedIndex].value);
   location.href=list.options[list.selectedIndex].value;
  }
  function Getseconditem(i,j)
  {//求大类的小类列表
   var unit_code;
   if(j==1)
     unit_code=document.input3.unit_code1.options[document.input3.unit_code1.selectedIndex].value;
   else
     {if(j==2)
        unit_code=document.input3.unit_code2.options[document.input3.unit_code2.selectedIndex].value; 
      else
        {if(j==3)
           unit_code=document.input3.unit_code3.options[document.input3.unit_code3.selectedIndex].value; 
        } 
     }
   //alert(i);
   location.href="searchfmc.asp?mode=1&unit_code="+unit_code;             
   return false;
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
  showchead()
  opendb()
  set rs1=server.createobject("adodb.recordset")
  set rsMX=server.createobject("adodb.recordset")
  if right(left(session("unit_code"),unit_len1),2)="00" then'省厅前ajlb_len1的后两位为0,即地区中后两位为0
    rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn, 1, 1  
  elseif right(session("unit_code"),unit_len0-unit_len2) = unit_str2 and mid(session("unit_code"),unit_len1+1,unit_len2-unit_len1)="00" then'市局
    rs1.open "select * from unit where unit_code='" & left(session("unit_code"),unit_len1) & "0000' order by unit_sxh", conn, 1, 1    
  else'单个收费站
    rs1.open "select * from unit where unit_code='" & left(session("unit_code"),unit_len1) & "0000' order by unit_sxh", conn, 1, 1    
  end if
  %>
  <form method="POST" action="searchfmc.asp?mode=2" name="input3">
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>花名册查询</b></td>
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
          <td align="center">
            地区
            <select name="unit_code1" style="HEIGHT:17px;WIDTH:59px" onchange="Getseconditem(301,1)">
              <%while not rs1.EOF 
                if trim(unit_code)="" then unit_code=trim(rs1("unit_code"))%>
                <option value="<%=trim(rs1("unit_code"))%>"<%if left(unit_code,unit_len1)=left(rs1("unit_code"),unit_len1) then %> selected <% end if %>><%=trim(rs1("unit_name"))%></option>
                <%rs1.MoveNext 
              WEND%>
            </select>
            收费站
            <select name="unit_code2" style="HEIGHT:17px;WIDTH:59px">
              <% 
              if right(left(session("unit_code"),unit_len1),2)="00" then'省厅前ajlb_len1的后两位为0,即地区中后两位为0
                rsMX.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn, 1, 1%>
                <option value="" <%if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then %> selected <% end if %>></option>
              <%elseif right(session("unit_code"),unit_len0-unit_len2) = unit_str2 and mid(session("unit_code"),unit_len1+1,unit_len2-unit_len1)="00" then'市局
                rsMX.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn, 1, 1%>
                <option value="" <%if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then %> selected <% end if %>></option>
              <%else'单个收费站
                rsMX.open "select * from unit where unit_code='" & left(session("unit_code"),unit_len2) & "00' order by unit_sxh", conn, 1, 1
              end if
              while not rsMX.EOF%>
                <option value="<%=trim(rsMX("unit_code"))%>"<%if left(unit_code,unit_len2)=left(rsMX("unit_code"),unit_len2) then %> selected <% end if %>><%=trim(rsMX("unit_name"))%></option>
                <%rsMX.MoveNext 
              WEND
              rsMX.close%>
            </select>
          </td>
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
  ShowInputForm3 ""
elseif mode=2 then
  if request("dq")="" then
    if request("unit_code2")="" then
      dq=request("unit_code1")
    else
      dq=request("unit_code2")
    end if
  else 
    dq=request("dq")
  end if
  'response.write dq
  'if dq <> "" then
    if not isEmpty(request("page")) and isnumeric(request("page")) then
      cpage = clng(request("page"))
    else
      cpage = 1
    end if
    opendb()
    set rs=server.createobject("adodb.recordset")
    set rs1=server.createobject("adodb.recordset")
    set rs2=server.createobject("adodb.recordset")
    if request("unit_code2")="" then
    '  if right(left(request("unit_code1"),unit_len1),2)="00" then'全省
    '    sql=" right(left(unit_code," & unit_len1 &"),2)<>'00' "
    '    sql1=" and right(left(unit_code," & unit_len1 &"),2)<>'00' "
    '    sql2=" and right(left(unit_code," & unit_len1 &"),2)<>'00' "
    '  else'全区
        sql=" unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
        sql1=" and unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
        sql2=" and unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
    '  end if
    else'收费站
      sql=" unit_code='" & request("unit_code2") & "'"
      sql1=" and unit_code='" & request("unit_code2") & "'"
      sql2=" and unit_code='" & request("unit_code2") & "'"
    end if
    bbjzdwmc=""
    bbjzzg=""
    bbjzfh=""
    bbjzzb=""
    'if request("unit_code2")="" then
    '  if right(left(request("unit_code1"),unit_len1),2)="00" then'全省
    '    rs.open "select * from unit where unit_code='" & request("unit_code1") & "'" ,conn,1,1
    '  else'全区
        rs.open "select * from unit where unit_code='" & request("unit_code1") & "'" ,conn,1,1
    '  end if
    'else'收费站
    '    rs.open "select * from unit where unit_code='" & request("unit_code2") & "'" ,conn,1,1
    'end if
    if rs.recordcount>0 then
      if not isnull(rs("bbjzdwmc")) then bbjzdwmc=rs("bbjzdwmc")
      if not isnull(rs("bbjzzg")) then bbjzzg=rs("bbjzzg")
      if not isnull(rs("bbjzfh")) then bbjzfh=rs("bbjzfh")
      if not isnull(rs("bbjzzb")) then bbjzzb=rs("bbjzzb")
    end if
    rs.close
    'response.write sql
    rs.open "select * from fmc where " + sql, conn, 1, 1
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
            [<a href="searchfmc.asp?mode=2">所有</a>][<a href="searchfmc.asp?mode=1">查找</a>]
          </td>
        </tr>
        <tr>
          <td valign="bottom">第<%=cstr(cpage)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个花名册记录</td>
          <td align="right">
            <%if cpage <> 1 then%>
              [<a href="searchfmc.asp?mode=2&unit_code2=<%=dq%>&page=<%=cstr(cpage-1)%>">上一页</a>]
            <%end if%>
            <%if cpage <> rs.PageCount then%>
              [<a href="searchfmc.asp?mode=2&unit_code2=<%=dq%>&page=<%=cstr(cpage+1)%>">下一页</a>]
            <%end if%>
            <%if rs.PageCount > 1 then%>
              <select name="select2"  onchange="goto(this)">
                <%for i = 1 to rs.PageCount%>
                  <%if i = cpage then%>
                    <option selected value="searchfmc.asp?mode=2&unit_code2=<%=dq%>&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
                  <%else%>
                    <option value="searchfmc.asp?mode=2&unit_code2=<%=dq%>&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
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
                <td align=center width=80>姓名</td>
                <td align=center width=60>性别</td>
				<td align=center width=80>出生年月</td>
				<td align=center width=80>政治面貌</td>
				<td align=center width=60>学历</td>
				<td align=center width=80>职务</td>
				<td width=100 align=center>收费证号</td>
                <td width=100 align=center>执法证号</td>
				<td align=center width=80>着装情况</td>
				<td align=center width=80>入伍时间</td>
				<td align=center width=120>工作单位</td>
				<td align=center width=200>奖惩及其他情况</td>
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
				  <td width=60>
                    <%if isnull(rs("xm")) then
					  response.write "&nbsp;"
					elseif rs("xm")<>"" then
                      response.write rs("xm")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
                  <td width=80>
                    <%if isnull(rs("xb")) then
					  response.write "&nbsp;"
					elseif rs("xb")<>"" then
                      response.write rs("xb")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
                  <td width=80>
                    <%if isnull(rs("csly")) then
					  response.write "&nbsp;"
					elseif rs("csly")<>"" then
                      response.write rs("csly")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=80>
                    <%if isnull(rs("dty")) then
					  response.write "&nbsp;"
					elseif rs("dty")<>"" then
                      response.write rs("dty")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=60>
                    <%if isnull(rs("wfcdxl")) then
					  response.write "&nbsp;"
					elseif rs("wfcdxl")<>"" then
                      response.write rs("wfcdxl")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=80>
                    <%if isnull(rs("zw")) then
					  response.write "&nbsp;"
					elseif rs("zw")<>"" then
                      response.write rs("zw")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=100>
                    <%if isnull(rs("sfzh")) then
					  response.write "&nbsp;"
					elseif rs("sfzh")<>"" then
                      response.write rs("sfzh")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=100>
                    <%if isnull(rs("zfzh")) then
					  response.write "&nbsp;"
					elseif rs("zfzh")<>"" then
                      response.write rs("zfzh")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=80>
                    <%if isnull(rs("zzqk")) then
					  response.write "&nbsp;"
					elseif rs("zzqk")<>"" then
                      response.write rs("zzqk")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=80>
                    <%if isnull(rs("rwly")) then
					  response.write "&nbsp;"
					elseif rs("rwly")<>"" then
                      response.write rs("rwly")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=120>
                    <%if isnull(rs("xdw")) then
					  response.write "&nbsp;"
					elseif rs("xdw")<>"" then
                      response.write rs("xdw")
                    else
                      response.write "&nbsp;"
                    end if%>
                  </td>
				  <td width=200>
                    <%if isnull(rs("jc")) then
					  response.write "&nbsp;"
					elseif rs("jc")<>"" then
                      response.write rs("jc")
                    else
                      response.write "&nbsp;"
                    end if%>
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
  'else
  '  ShowInputForm3 ""
  'end if
end if
%>    