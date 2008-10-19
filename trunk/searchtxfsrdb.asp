<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<!--#include file="./fcommon.asp"-->

<%
if session("username")=""  or instr(session("power"),",3,")=0 then
  Response.Redirect("notlogin.asp")
end if

dim conn, rs, rs1,rsMX,rs2,rs3,rs4, sql,sql1,sql2,sql3,sql4,sql5, errmsg, founderror, i, str1, mode, cpage, fl,dwx,unit_code,shj1,shj2,shj3,byhj,qnhj,bnlj,qnlj,bbjzdwmc,bbjzzg,bbjzfh,bbjzzb

if not isempty(request("mode")) and isnumeric(request("mode")) then
    mode = clng(request("mode"))
else
    mode=1
end if
if not isempty(request("unit_code")) then
    unit_code = request("unit_code")
else
    unit_code = ""
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
  <title>公路通行费收入对比表</title>
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
  <form method="POST" action="searchtxfsrdb.asp?mode=1" name="input3">
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>通行费收入对比表查询</b></td>
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
            <select name="unit_code1" style="HEIGHT:17px;WIDTH:59px">
              <%while not rs1.EOF 
                if trim(unit_code)="" then unit_code=trim(rs1("unit_code"))%>
                <option value="<%=trim(rs1("unit_code"))%>"<%if left(unit_code,unit_len1)=left(rs1("unit_code"),unit_len1) then %> selected <% end if %>><%=trim(rs1("unit_name"))%></option>
                <%rs1.MoveNext 
              WEND%>
            </select>
            &nbsp;&nbsp;
            <select name="shj1" style="HEIGHT:17px;WIDTH:50px" >
              <%
              for i=2004 to year(now)%>
                <option value="<%=i%>"<%if shj1=year(now) then %> selected <% end if %>><%=i%></option>
                <%
              next
              %>
            </select>
            年
            <select name="shj2" style="HEIGHT:17px;WIDTH:50px" >
              <option value="<%=0%>"<%if shj2=0 then %> selected <% end if %>><%=""%></option>
              <%
              for i=1 to 12%>
                <%if i<10 then%>
                  <option value="<%="0"&i%>"<%if shj2=month(now) then %> selected <% end if %>><%="0"&i%></option>
                <%else%>
                  <option value="<%=i%>"<%if shj2=month(now) then %> selected <% end if %>><%=i%></option>
                <%end if%>
                <%
              next
              %>
            </select>
            月
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
  if trim(request("shj2")) <> "0" and trim(request("shj2"))<>"" then
    if not isEmpty(request("page")) and isnumeric(request("page")) then
      cpage = clng(request("page"))
    else
      cpage = 1
    end if
    opendb()
    set rs=server.createobject("adodb.recordset")
    set rs1=server.createobject("adodb.recordset")
    set rs2=server.createobject("adodb.recordset")
    set rs3=server.createobject("adodb.recordset")
    set rs4=server.createobject("adodb.recordset")
    if right(left(request("unit_code1"),unit_len1),2)="00" then'全省
      sql=" right(left(unit_code," & unit_len1 &"),2)<>'00' "
      sql1=" and right(left(unit.unit_code," & unit_len1 &"),2)<>'00' "
      sql2=" and right(left(unit.unit_code," & unit_len1 &"),2)<>'00' "
      sql3=" and right(left(unit.unit_code," & unit_len1 &"),2)<>'00' "
      sql4=" and right(left(unit.unit_code," & unit_len1 &"),2)<>'00' "
    else'全区
      sql=" left(unit_code," & unit_len1 &")='" & left(request("unit_code1"),unit_len1) & "' and mid(unit_code,"&unit_len1+1&","&unit_len0-unit_len2&")<>'00'"
      sql1=" and left(unit.unit_code," & unit_len1 &")='" & left(request("unit_code1"),unit_len1) & "' and mid(unit.unit_code,"&unit_len1+1&","&unit_len0-unit_len2&")<>'00'"
      sql2=" and left(unit.unit_code," & unit_len1 &")='" & left(request("unit_code1"),unit_len1) & "' and mid(unit.unit_code,"&unit_len1+1&","&unit_len0-unit_len2&")<>'00'"
      sql3=" and left(unit.unit_code," & unit_len1 &")='" & left(request("unit_code1"),unit_len1) & "' and mid(unit.unit_code,"&unit_len1+1&","&unit_len0-unit_len2&")<>'00'"
      sql4=" and left(unit.unit_code," & unit_len1 &")='" & left(request("unit_code1"),unit_len1) & "' and mid(unit.unit_code,"&unit_len1+1&","&unit_len0-unit_len2&")<>'00'"
    end if
    bbjzdwmc=""
    bbjzzg=""
    bbjzfh=""
    bbjzzb=""
    if right(left(request("unit_code1"),unit_len1),2)="00" then'全省
      rs.open "select * from unit where unit_code='" & request("unit_code1") & "'" ,conn,1,1
    else'全区
      rs.open "select * from unit where unit_code='" & request("unit_code1") & "'" ,conn,1,1
    end if
    if rs.recordcount>0 then
      if not isnull(rs("bbjzdwmc")) then bbjzdwmc=rs("bbjzdwmc")
      if not isnull(rs("bbjzzg")) then bbjzzg=rs("bbjzzg")
      if not isnull(rs("bbjzfh")) then bbjzfh=rs("bbjzfh")
      if not isnull(rs("bbjzzb")) then bbjzzb=rs("bbjzzb")
    end if
    rs.close
    shj3=request("shj1")-1
    if trim(request("shj2")) <> "0" and trim(request("shj2"))<>"" then
      if sql<>"" then sql=sql & " and "
      if sql1<>"" then sql1=sql1 & " and "
      if sql2<>"" then sql2=sql2 & " and "
      if sql3<>"" then sql3=sql3 & " and "
      if sql4<>"" then sql4=sql4 & " and "
      sql=sql & " (shj1 like '" + trim(request("shj1")) +trim(request("shj2")) + "%')"
      sql1=sql1 & " (shj1 like '" + trim(shj3) +trim(request("shj2")) + "%')"
      sql2=sql2 & " (shj1 like '" + trim(shj3) + "%' and left(shj1,6)<='" + trim(request("shj1")) + trim(request("shj2")) +"')"
      sql3=sql4 & " (shj1 like '" + trim(request("shj1")) +trim(request("shj2")) + "%')"
      sql4=sql4 & " (shj1 like '" + trim(request("shj1")) + "%' and left(shj1,6)<='" + trim(request("shj1")) + trim(request("shj2")) +"')"
    end if
    'response.write sql1
    rs.open "select * from edzdjb03 where " + sql, conn, 1, 1
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
            [<a href="searchtxfsrdb.asp?mode=1">继续查找</a>]
          </td>
        </tr>
        <tr>
          <td align="center" colspan="2">
            福建省普通公路通行费收入对比表
          </td>
        </tr>
        <tr>
          <td align="center" colspan="2">
            <%=bbjzdwmc & "-" & request("shj1")&"年" & request("shj2")& "月"%>
          </td>
        </tr>
        <tr>
          <td align="left">
            编报单位（盖章）
          </td>
          <td align="right">
            报送日期：<%=year(now)%>年<%=month(now)%>月<%=day(now)%>日
          </td>
        </tr>
        <tr>
          <td colspan="2">
            <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <tr bgcolor=<%=skincolor()%>>
                <td align=center rowspan="2">单位</td>
                <td align=center colspan="2">本月数</td>
                <td align=center colspan="2">累计数</td>
              </tr>
              <tr bgcolor=<%=skincolor()%>>
                <td width=140 align=center>去年同期</td>
                <td width=140 align=center>本月收入</td>
                <td width=140 align=center>去年累计</td>
                <td width=140 align=center>本年累计</td>
              </tr>
              <%
              byhj=0
              qnhj=0
              bnlj=0
              qnlj=0
              fl=true
              rs.close
              if right(left(request("unit_code1"),unit_len1),2)="00" then'全省各区比较
                rs.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' and right(left(unit.unit_code," & unit_len1 &"),2)<>'00' order by unit_sxh", conn, 1, 1 
              else'全区各收费站比较
                rs.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(request("unit_code1"),unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn, 1, 1
              end if
              if rs.recordcount>0 then
                sql5=" left(ajlb_code," & ajlb_len2 & ")='" & left("0301000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00'"
                sql5=sql5 & " or left(ajlb_code," & ajlb_len2 & ")='" & left("0302000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00'"
                sql5=sql5 & " or ajlb_code='" & "0303000000" &"' or ajlb_code='" & "0304000000" &"'"
                if sql5<>"" then sql5=" and ("& sql5 & ") "
                'response.write "SELECT left(unit.unit_code,6) as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code and left(unit.unit_code,4)='3504' and mid(unit.unit_code,5,2)<>'00' and (shj1 like '200504%') GROUP BY left(unit.unit_code,6)"
                if right(left(request("unit_code1"),unit_len1),2)="00" then'全省各区比较
                  rs1.open "SELECT left(unit.unit_code,"& unit_len1 & ") as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code " & sql1 & sql5 & " GROUP BY left(unit.unit_code,"& unit_len1 & ")",conn,1,1
                  rs2.open "SELECT left(unit.unit_code,"& unit_len1 & ") as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code " & sql2 & sql5 & " GROUP BY left(unit.unit_code,"& unit_len1 & ")",conn,1,1
                  rs3.open "SELECT left(unit.unit_code,"& unit_len1 & ") as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code " & sql3 & sql5 & " GROUP BY left(unit.unit_code,"& unit_len1 & ")",conn,1,1
                  rs4.open "SELECT left(unit.unit_code,"& unit_len1 & ") as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code " & sql4 & sql5 & " GROUP BY left(unit.unit_code,"& unit_len1 & ")",conn,1,1
                else'全区各收费站比较
                  rs1.open "SELECT left(unit.unit_code,"& unit_len2 & ") as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code " & sql1 & sql5 & " GROUP BY left(unit.unit_code,"& unit_len2 & ")",conn,1,1
                  rs2.open "SELECT left(unit.unit_code,"& unit_len2 & ") as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code " & sql2 & sql5 & " GROUP BY left(unit.unit_code,"& unit_len2 & ")",conn,1,1
                  rs3.open "SELECT left(unit.unit_code,"& unit_len2 & ") as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code " & sql3 & sql5 & " GROUP BY left(unit.unit_code,"& unit_len2 & ")",conn,1,1
                  rs4.open "SELECT left(unit.unit_code,"& unit_len2 & ") as expr2,sum(edzdjb_x03.ajlbV) as expr1 FROM edzdjb03,edzdjb_x03,unit WHERE edzdjb03.bh=edzdjb_x03.bh and edzdjb03.unit_code=unit.unit_code " & sql4 & sql5 & " GROUP BY left(unit.unit_code,"& unit_len2 & ")",conn,1,1
                end if
                response.write "<tr>"
                do while not rs.eof 
                  if not fl then
                    response.write "<tr>"
                    fl=true
                  end if
                  response.write "<td>" & rs("unit_name") & "</td>"
                  '去年同期
                  if rs1.recordcount>0 then
                    rs1.movefirst
                    if right(left(request("unit_code1"),unit_len1),2)="00" then'全省各区比较
                      rs1.find "expr2='" & left(rs("unit_code"),unit_len1) & "'"
                    else'全区各收费站比较
                      rs1.find "expr2='" & left(rs("unit_code"),unit_len2) & "'"
                    end if
                    if not rs1.eof then
                      qnhj=qnhj+rs1("expr1")
                      response.write "<td>" & rs1("expr1") & "</td>"
                    else
                      response.write "<td>0</td>"
                    end if
                  else
                    response.write "<td>0</td>"
                  end if
                  '本月收入
                  if rs3.recordcount>0 then
                    rs3.movefirst
                    if right(left(request("unit_code1"),unit_len1),2)="00" then'全省各区比较
                      rs3.find "expr2='" & left(rs("unit_code"),unit_len1) & "'"
                    else'全区各收费站比较
                      rs3.find "expr2='" & left(rs("unit_code"),unit_len2) & "'"
                    end if
                    if not rs3.eof then
                      byhj=byhj+rs3("expr1")
                      response.write "<td>" & rs3("expr1") & "</td>"
                    else
                      response.write "<td>0</td>"
                    end if
                  else
                    response.write "<td>0</td>"
                  end if
                  '去年累计
                  if rs2.recordcount>0 then
                    rs2.movefirst
                    if right(left(request("unit_code1"),unit_len1),2)="00" then'全省各区比较
                      rs2.find "expr2='" & left(rs("unit_code"),unit_len1) & "'"
                    else'全区各收费站比较
                      rs2.find "expr2='" & left(rs("unit_code"),unit_len2) & "'"
                    end if
                    if not rs2.eof then
                      qnlj=qnlj+rs2("expr1")
                      response.write "<td>" & rs2("expr1") & "</td>"
                    else
                      response.write "<td>0</td>"
                    end if
                  else
                    response.write "<td>0</td>"
                  end if
                  '本年累计
                  if rs4.recordcount>0 then
                    rs4.movefirst
                    if right(left(request("unit_code1"),unit_len1),2)="00" then'全省各区比较
                      rs4.find "expr2='" & left(rs("unit_code"),unit_len1) & "'"
                    else'全区各收费站比较
                      rs4.find "expr2='" & left(rs("unit_code"),unit_len2) & "'"
                    end if
                    if not rs4.eof then
                      bnlj=bnlj+rs4("expr1")
                      response.write "<td>" & rs4("expr1") & "</td>"
                    else
                      response.write "<td>0</td>"
                    end if
                  else
                    response.write "<td>0</td>"
                  end if
                  response.write "</tr>"
                  fl=false
                  rs.movenext
                loop
                rs2.close
                rs1.close
              end if
              rs.close
              response.write "<tr>"
              response.write "<td align=center>合计</td>"
              response.write "<td align=center>" &qnhj & "</td>"
              response.write "<td align=center>" &byhj & "</td>"
              response.write "<td align=center>" &qnlj & "</td>"
              response.write "<td align=center>" &bnlj & "</td>"
              response.write "</tr>"
              %>
            </table>
          </td>
        </tr>
        <tr>
          <td align="center" colspan="2">
            <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <tr>
                <td align=left>主管：<%=bbjzzg%></td>
                <td align=center>复核：<%=bbjzfh%></td>
                <td align=right>制表：<%=bbjzzb%></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <%
      set rs=nothing
      closedb()
      showctail()
    end if
  else
    ShowInputForm3 ""
  end if
end if
%>    