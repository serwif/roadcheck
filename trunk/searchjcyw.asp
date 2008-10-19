<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<!--#include file="./fcommon.asp"-->

<%
if session("username")=""  or instr(session("power"),",3,")=0 then
  Response.Redirect("notlogin.asp")
end if

dim conn, rs, rs1,rsMX,rs2,rs3,rs4,rs5, sql,sql1,sql2,sql3,sql4,sql5, errmsg, founderror, i, str1, mode, cpage, fl,flbz,flbz1,dwx,unit_code,shj1,shj2,byhj,byhj2,bnlj,bnlj2,bbjzdwmc,bbjzzg,bbjzfh,bbjzzb,bzhs

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
  <title>稽查业务情况表</title>
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
   location.href="searchjcyw.asp?mode=1&unit_code="+unit_code;             
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
  <form method="POST" action="searchjcyw.asp?mode=1" name="input3">
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>通行费稽查业务情况表查询</b></td>
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
              for i=1 to 4%>
                <%if i<10 then%>
                  <option value="<%="0"&i%>"<%if shj2=month(now) then %> selected <% end if %>><%="0"&i%></option>
                <%else%>
                  <option value="<%=i%>"<%if shj2=month(now) then %> selected <% end if %>><%=i%></option>
                <%end if%>
                <%
              next
              %>
            </select>
            季度
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
    set rs5=server.createobject("adodb.recordset")
    if request("unit_code2")="" then
      if right(left(request("unit_code1"),unit_len1),2)="00" then'全省
        sql=" right(left(unit_code," & unit_len1 &"),2)<>'00' "
        sql1=" and right(left(unit_code," & unit_len1 &"),2)<>'00' "
        sql2=" and right(left(unit_code," & unit_len1 &"),2)<>'00' "
      else'全区
        sql=" unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
        sql1=" and unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
        sql2=" and unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
      end if
    else'收费站
      sql=" unit_code='" & request("unit_code2") & "'"
      sql1=" and unit_code='" & request("unit_code2") & "'"
      sql2=" and unit_code='" & request("unit_code2") & "'"
    end if
    bbjzdwmc=""
    bbjzzg=""
    bbjzfh=""
    bbjzzb=""
    if request("unit_code2")="" then
      if right(left(request("unit_code1"),unit_len1),2)="00" then'全省
        rs.open "select * from unit where unit_code='" & request("unit_code1") & "'" ,conn,1,1
      else'全区
        rs.open "select * from unit where unit_code='" & request("unit_code1") & "'" ,conn,1,1
      end if
    else'收费站
        rs.open "select * from unit where unit_code='" & request("unit_code2") & "'" ,conn,1,1
    end if
    if rs.recordcount>0 then
      if not isnull(rs("bbjzdwmc")) then bbjzdwmc=rs("bbjzdwmc")
      if not isnull(rs("bbjzzg")) then bbjzzg=rs("bbjzzg")
      if not isnull(rs("bbjzfh")) then bbjzfh=rs("bbjzfh")
      if not isnull(rs("bbjzzb")) then bbjzzb=rs("bbjzzb")
    end if
    rs.close
    if trim(request("shj2")) <> "0" and trim(request("shj2"))<>"" then
      if sql<>"" then sql=sql & " and "
      if sql1<>"" then sql1=sql1 & " and "
      if sql2<>"" then sql2=sql2 & " and "
      if request("shj2")=1 then
        sql=sql & " (left(shj1,6)>='" + trim(request("shj1")) + "01' and left(shj1,6)<='" + trim(request("shj1")) + "03')"
        sql1=sql1 & " (left(shj1,6)>='" + trim(request("shj1")) + "01' and left(shj1,6)<='" + trim(request("shj1")) + "03')"
      elseif request("shj2")=2 then
        sql=sql & " (left(shj1,6)>='" + trim(request("shj1")) + "04' and left(shj1,6)<='" + trim(request("shj1")) + "06')"
        sql1=sql1 & " (left(shj1,6)>='" + trim(request("shj1")) + "04' and left(shj1,6)<='" + trim(request("shj1")) + "06')"
      elseif request("shj2")=3 then
        sql=sql & " (left(shj1,6)>='" + trim(request("shj1")) + "07' and left(shj1,6)<='" + trim(request("shj1")) + "09')"
        sql1=sql1 & " (left(shj1,6)>='" + trim(request("shj1")) + "07' and left(shj1,6)<='" + trim(request("shj1")) + "09')"
      elseif request("shj2")=4 then
        sql=sql & " (left(shj1,6)>='" + trim(request("shj1")) + "10' and left(shj1,6)<='" + trim(request("shj1")) + "12')"
        sql1=sql1 & " (left(shj1,6)>='" + trim(request("shj1")) + "10' and left(shj1,6)<='" + trim(request("shj1")) + "12')"
      end if
      sql2=sql2 & " (shj1 like '" + trim(request("shj1")) + "%' and left(shj1,6)<='" + trim(request("shj1")) + trim(request("shj2")) +"')"
    end if
    'response.write sql
    rs.open "select * from edzdjb04 where " + sql, conn, 1, 1
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
            [<a href="searchjcyw.asp?mode=1">继续查找</a>]
          </td>
        </tr>
        <tr>
          <td align="center" colspan="2">
            福建省普通公路通行费稽查业务情况表
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
              <%
              byhj=0
              bnlj=0
              fl=false
              flbz=false
              flbz1=false
              bzhs=0'备注行数
              rs.close
              rs5.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0405000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1'违章车辆处理情况记录集
              if rs5.recordcount>0 then
                rs3.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs5("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                rs4.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs5("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql2 & " group by ajlb.ajlb_code", conn, 1, 1          
              end if
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0401000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              bzhs=bzhs+rs.recordcount+1
              rs.close
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0402000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              bzhs=bzhs+rs.recordcount+3
              rs.close
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0403000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              bzhs=bzhs+rs.recordcount+2
              rs.close
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0404000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              bzhs=bzhs+rs.recordcount+4
              rs.close
              bzhs=bzhs-(rs5.recordcount+3)
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0401000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                'response.write "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code"
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                rs2.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql2 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<tr>"
                response.write "<td align=center colspan=3>一.征管人员情况</td>"
                response.write "<td align=center colspan=3>五.违章车辆处理情况</td>"
                response.write "</tr>"
                do while not rs.eof 
                  response.write "<tr>"
                  response.write "<td align=center>" & rs("ajlb_name") & "</td>"
                  if rs1.recordcount>0 then
                    rs1.movefirst
                    rs1.find "ajlb_code='" & rs("ajlb_code") & "'"
                    if not rs1.eof then
                      byhj=byhj+rs1("expr1")
                      response.write "<td colspan=2 align=right>" & rs1("expr1") & "</td>"
                    else
                      response.write "<td colspan=2 align=right>0</td>"
                    end if
                  else
                    response.write "<td colspan=2 align=right>0</td>"
                  end if
                  if not fl then
                    response.write "<td align=center>项目</td>"
                    response.write "<td align=center>本季</td>"
                    response.write "<td align=center>年累</td>"
                    fl=true
                  else
                    if not rs5.eof then
                      response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                      if rs3.recordcount>0 then
                        rs3.movefirst
                        rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs3.eof then
                          response.write "<td align=right>" & rs3("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      if rs4.recordcount>0 then
                        rs4.movefirst
                        rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs4.eof then
                          response.write "<td align=right>" & rs4("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      rs5.movenext
                    else
                      if not flbz then
                        response.write "<td align=center colspan=3>备注</td>"
                        flbz=true          
                      elseif not flbz1 then
                        if bzhs>0 then
                          response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                        end if
                        flbz1=true                      
                      end if
                    end if
                  end if
                  response.write "</tr>"
                  rs.movenext
                loop
                rs2.close
                rs1.close
              end if
              rs.close
              byhj=0
              bnlj=0
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0402000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                rs2.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql2 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<tr>"
                response.write "<td align=center colspan=3>二.稽查情况</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
                response.write "<tr>"
                response.write "<td align=center>项目</td>"
                response.write "<td align=center>本季</td>"
                response.write "<td align=center>年累计</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
                do while not rs.eof 
                  response.write "<tr>"
                  response.write "<td align=center>" & rs("ajlb_name") & "</td>"
                  if rs1.recordcount>0 then
                    rs1.movefirst
                    rs1.find "ajlb_code='" & rs("ajlb_code") & "'"
                    if not rs1.eof then
                      byhj=byhj+rs1("expr1")
                      response.write "<td align=right>" & rs1("expr1") & "</td>"
                    else
                      response.write "<td align=right>0</td>"
                    end if
                  else
                    response.write "<td align=right>0</td>"
                  end if
                  if rs2.recordcount>0 then
                    rs2.movefirst
                    rs2.find "ajlb_code='" & rs("ajlb_code") & "'"
                    if not rs2.eof then
                      bnlj=bnlj+rs2("expr1")
                      response.write "<td align=right>" & rs2("expr1") & "</td>"
                    else
                      response.write "<td align=right>0</td>"
                    end if
                  else
                    response.write "<td align=right>0</td>"
                  end if
                  if not fl then
                    response.write "<td align=center>项目</td>"
                    response.write "<td align=center>本季</td>"
                    response.write "<td align=center>年累</td>"
                    fl=true
                  else
                    if not rs5.eof then
                      response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                      if rs3.recordcount>0 then
                        rs3.movefirst
                        rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs3.eof then
                          response.write "<td align=right>" & rs3("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      if rs4.recordcount>0 then
                        rs4.movefirst
                        rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs4.eof then
                          response.write "<td align=right>" & rs4("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      rs5.movenext
                    else
                      if not flbz then
                        response.write "<td align=center colspan=3>备注</td>"
                        flbz=true          
                      elseif not flbz1 then
                        if bzhs>0 then
                          response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                        end if
                        flbz1=true                      
                      end if
                    end if
                  end if
                  response.write "</tr>"
                  rs.movenext
                loop
                rs2.close
                rs1.close
                response.write "<tr>"
                response.write "<td align=center>内查外查合计数</td>"
                response.write "<td align=right>"&byhj&"</td>"
                response.write "<td align=right>"&bnlj&"</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
              end if
              rs.close
              byhj=0
              bnlj=0
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0403000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' and ajlb_code<>'0403030000' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                rs2.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql2 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<tr>"
                response.write "<td align=center colspan=3>三.征管人员违章违纪情况</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
                do while not rs.eof 
                  response.write "<tr>"
                  response.write "<td align=center>" & rs("ajlb_name") & "</td>"
                  if rs1.recordcount>0 then
                    rs1.movefirst
                    rs1.find "ajlb_code='" & rs("ajlb_code") & "'"
                    if not rs1.eof then
                      byhj=byhj+rs1("expr1")
                      response.write "<td align=right>" & rs1("expr1") & "</td>"
                    else
                      response.write "<td align=right>0</td>"
                    end if
                  else
                    response.write "<td align=right>0</td>"
                  end if
                  if rs2.recordcount>0 then
                    rs2.movefirst
                    rs2.find "ajlb_code='" & rs("ajlb_code") & "'"
                    if not rs2.eof then
                      bnlj=bnlj+rs2("expr1")
                      response.write "<td align=right>" & rs2("expr1") & "</td>"
                    else
                      response.write "<td align=right>0</td>"
                    end if
                  else
                    response.write "<td align=right>0</td>"
                  end if
                  if not fl then
                    response.write "<td align=center>项目</td>"
                    response.write "<td align=center>本季</td>"
                    response.write "<td align=center>年累</td>"
                    fl=true
                  else
                    if not rs5.eof then
                      response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                      if rs3.recordcount>0 then
                        rs3.movefirst
                        rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs3.eof then
                          response.write "<td align=right>" & rs3("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      if rs4.recordcount>0 then
                        rs4.movefirst
                        rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs4.eof then
                          response.write "<td align=right>" & rs4("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      rs5.movenext
                    else
                      if not flbz then
                        response.write "<td align=center colspan=3>备注</td>"
                        flbz=true          
                      elseif not flbz1 then
                        if bzhs>0 then
                          response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                        end if
                        flbz1=true                      
                      end if
                    end if
                  end if
                  response.write "</tr>"
                  rs.movenext
                loop
                rs2.close
                rs1.close
                response.write "<tr>"
                response.write "<td align=center>合计(人数)</td>"
                response.write "<td align=right>"&byhj&"</td>"
                response.write "<td align=right>"&bnlj&"</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
              end if
              rs.close
              byhj=0
              bnlj=0
              rs.open "select * from ajlb where ajlb_code='0403030000' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                rs2.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql2 & " group by ajlb.ajlb_code", conn, 1, 1
                do while not rs.eof 
                  response.write "<tr>"
                  response.write "<td align=center>" & rs("ajlb_name") & "</td>"
                  if rs1.recordcount>0 then
                    rs1.movefirst
                    rs1.find "ajlb_code='" & rs("ajlb_code") & "'"
                    if not rs1.eof then
                      byhj=byhj+rs1("expr1")
                      response.write "<td align=right>" & rs1("expr1") & "</td>"
                    else
                      response.write "<td align=right>0</td>"
                    end if
                  else
                    response.write "<td align=right>0</td>"
                  end if
                  if rs2.recordcount>0 then
                    rs2.movefirst
                    rs2.find "ajlb_code='" & rs("ajlb_code") & "'"
                    if not rs2.eof then
                      bnlj=bnlj+rs2("expr1")
                      response.write "<td align=right>" & rs2("expr1") & "</td>"
                    else
                      response.write "<td align=right>0</td>"
                    end if
                  else
                    response.write "<td align=right>0</td>"
                  end if
                  if not fl then
                    response.write "<td align=center>项目</td>"
                    response.write "<td align=center>本季</td>"
                    response.write "<td align=center>年累</td>"
                    fl=true
                  else
                    if not rs5.eof then
                      response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                      if rs3.recordcount>0 then
                        rs3.movefirst
                        rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs3.eof then
                          response.write "<td align=right>" & rs3("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      if rs4.recordcount>0 then
                        rs4.movefirst
                        rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs4.eof then
                          response.write "<td align=right>" & rs4("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      rs5.movenext
                    else
                      if not flbz then
                        response.write "<td align=center colspan=3>备注</td>"
                        flbz=true          
                      elseif not flbz1 then
                        if bzhs>0 then
                          response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                        end if
                        flbz1=true                      
                      end if
                    end if
                  end if
                  response.write "</tr>"
                  rs.movenext
                loop
                rs2.close
                rs1.close
              end if
              rs.close
              byhj=0
              byhj2=0
              bnlj=0
              bnlj2=0
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0404000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                'response.write "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code"
                'rs1.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                'rs2.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql2 & " group by ajlb.ajlb_code", conn, 1, 1
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                rs2.open "select ajlb.ajlb_code,sum(edzdjb_x04.ajlbV) as expr1 from edzdjb04,edzdjb_x04,ajlb where edzdjb04.bh=edzdjb_x04.bh and edzdjb_x04.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql2 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<tr>"
                response.write "<td align=center colspan=3>四.征费设施损坏及赔偿情况</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
                response.write "<tr>"
                response.write "<td align=center>损坏设施名称</td>"
                response.write "<td align=center>次数</td>"
                response.write "<td align=center>赔偿金额</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
                do while not rs.eof 
                  response.write "<tr>"
                  response.write "<td align=center>" & rs("ajlb_name") & "</td>"
                  if rs1.recordcount>0 then
                    rs1.movefirst
                    rs1.find "ajlb_code='" & left(rs("ajlb_code"),ajlb_len3) & "0100'"
                    if not rs1.eof then
                      byhj=byhj+rs1("expr1")
                      response.write "<td align=right>" & rs1("expr1") & "</td>"
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs1.movefirst
                    rs1.find "ajlb_code='" & left(rs("ajlb_code"),ajlb_len3) & "0200'"
                    if not rs1.eof then
                      byhj2=byhj2+rs1("expr1")
                      response.write "<td align=right>" & rs1("expr1") & "</td>"
                    else
                      response.write "<td align=right>0</td>"
                    end if
                  else
                    response.write "<td align=right>0</td>"
                    response.write "<td align=right>0</td>"
                  end if
                  if rs2.recordcount>0 then
                    rs2.movefirst
                    rs2.find "ajlb_code='" & left(rs("ajlb_code"),ajlb_len3) & "0100'"
                    if not rs2.eof then
                      bnlj=bnlj+rs2("expr1")
                    else
                    end if
                    rs2.movefirst
                    rs2.find "ajlb_code='" & left(rs("ajlb_code"),ajlb_len3) & "0200'"
                    if not rs2.eof then
                      bnlj2=bnlj2+rs2("expr1")
                    else
                    end if
                  else
                  end if
                  if not fl then
                    response.write "<td align=center>项目</td>"
                    response.write "<td align=center>本季</td>"
                    response.write "<td align=center>年累</td>"
                    fl=true
                  else
                    if not rs5.eof then
                      response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                      if rs3.recordcount>0 then
                        rs3.movefirst
                        rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs3.eof then
                          response.write "<td align=right>" & rs3("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      if rs4.recordcount>0 then
                        rs4.movefirst
                        rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                        if not rs4.eof then
                          response.write "<td align=right>" & rs4("expr1") & "</td>"
                        else
                          response.write "<td align=right>0</td>"
                        end if
                      else
                        response.write "<td align=right>0</td>"
                      end if
                      rs5.movenext
                    else
                      if not flbz then
                        response.write "<td align=center colspan=3>备注</td>"
                        flbz=true          
                      elseif not flbz1 then
                        if bzhs>0 then
                          response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                        end if
                        flbz1=true                      
                      end if
                    end if
                  end if
                  response.write "</tr>"
                  rs.movenext
                loop
                rs2.close
                rs1.close
                response.write "<tr>"
                response.write "<td align=center>合计</td>"
                response.write "<td align=right>"&byhj&"</td>"
                response.write "<td align=right>"&byhj2&"</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
                response.write "<tr>"
                response.write "<td align=center>年累计</td>"
                response.write "<td align=right>"&bnlj&"</td>"
                response.write "<td align=right>"&bnlj2&"</td>"
                if not fl then
                  response.write "<td align=center>项目</td>"
                  response.write "<td align=center>本季</td>"
                  response.write "<td align=center>年累</td>"
                  fl=true
                else
                  if not rs5.eof then
                    response.write "<td align=center>" & rs5("ajlb_name") & "</td>"
                    if rs3.recordcount>0 then
                      rs3.movefirst
                      rs3.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs3.eof then
                        response.write "<td align=right>" & rs3("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    if rs4.recordcount>0 then
                      rs4.movefirst
                      rs4.find "ajlb_code='" & rs5("ajlb_code") & "'"
                      if not rs4.eof then
                        response.write "<td align=right>" & rs4("expr1") & "</td>"
                      else
                        response.write "<td align=right>0</td>"
                      end if
                    else
                      response.write "<td align=right>0</td>"
                    end if
                    rs5.movenext
                  else
                    if not flbz then
                      response.write "<td align=center colspan=3>备注</td>"
                      flbz=true          
                    elseif not flbz1 then
                      if bzhs>0 then
                        response.write "<td align=center rowspan=" & bzhs & " colspan=3>&nbsp;</td>"
                      end if
                      flbz1=true                      
                    end if
                  end if
                end if
                response.write "</tr>"
              end if
              rs.close
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