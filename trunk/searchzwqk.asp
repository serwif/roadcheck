<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<!--#include file="./fcommon.asp"-->

<%
if session("username")=""  or instr(session("power"),",3,")=0 then
  Response.Redirect("notlogin.asp")
end if

dim conn, rs, rs1,rsMX,rs2,rs3, sql,sql1,sql2,sql3,sql4,sql5, errmsg, founderror, i, str1, mode, cpage, fl,dwx,unit_code,shj1,shj2,byhj,bnlj,bbjzdwmc,bbjzzg,bbjzfh,bbjzzb,xxhs,fdhs

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
  <title>ծ�������</title>
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
  {//������С���б�
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
   location.href="searchzwqk.asp?mode=1&unit_code="+unit_code;             
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
  if right(left(session("unit_code"),unit_len1),2)="00" then'ʡ��ǰajlb_len1�ĺ���λΪ0,�������к���λΪ0
    rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn, 1, 1  
  elseif right(session("unit_code"),unit_len0-unit_len2) = unit_str2 and mid(session("unit_code"),unit_len1+1,unit_len2-unit_len1)="00" then'�о�
    rs1.open "select * from unit where unit_code='" & left(session("unit_code"),unit_len1) & "0000' order by unit_sxh", conn, 1, 1    
  else'�����շ�վ
    rs1.open "select * from unit where unit_code='" & left(session("unit_code"),unit_len1) & "0000' order by unit_sxh", conn, 1, 1    
  end if
  %>
  <form method="POST" action="searchzwqk.asp?mode=1" name="input3">
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>ծ��������ѯ</b></td>
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
          <td align="center">
            ����
            <select name="unit_code1" style="HEIGHT:17px;WIDTH:59px" onchange="Getseconditem(301,1)">
              <%while not rs1.EOF 
                if trim(unit_code)="" then unit_code=trim(rs1("unit_code"))%>
                <option value="<%=trim(rs1("unit_code"))%>"<%if left(unit_code,unit_len1)=left(rs1("unit_code"),unit_len1) then %> selected <% end if %>><%=trim(rs1("unit_name"))%></option>
                <%rs1.MoveNext 
              WEND%>
            </select>
            �շ�վ
            <select name="unit_code2" style="HEIGHT:17px;WIDTH:59px">
              <% 
              if right(left(session("unit_code"),unit_len1),2)="00" then'ʡ��ǰajlb_len1�ĺ���λΪ0,�������к���λΪ0
                rsMX.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn, 1, 1%>
                <option value="" <%if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then %> selected <% end if %>></option>
              <%elseif right(session("unit_code"),unit_len0-unit_len2) = unit_str2 and mid(session("unit_code"),unit_len1+1,unit_len2-unit_len1)="00" then'�о�
                rsMX.open "select * from unit where left(unit_code," & unit_len1 & ")='" & left(unit_code,unit_len1) &"' and right(unit_code,"& (unit_len0-unit_len2) & ")='" & unit_str2 &"' and mid(unit_code,"& (unit_len1+1) & "," & (unit_len2-unit_len1) & ")<>'00' order by unit_sxh", conn, 1, 1%>
                <option value="" <%if mid(unit_code,unit_len1+1,unit_len2-unit_len1)="00" then %> selected <% end if %>></option>
              <%else'�����շ�վ
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
              <option value="<%=0%>"<%if shj1=0 then %> selected <% end if %>><%=""%></option>
              <%
              for i=2004 to year(now)%>
                <option value="<%=i%>"<%if shj1=year(now) then %> selected <% end if %>><%=i%></option>
                <%
              next
              %>
            </select>
            ��
          </td>
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
  '����
  if trim(request("shj1")) <> "0" and trim(request("shj1"))<>"" then
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
      if right(left(request("unit_code1"),unit_len1),2)="00" then'ȫʡ
        sql=" right(left(unit_code," & unit_len1 &"),2)<>'00' "
        sql1=" and right(left(unit_code," & unit_len1 &"),2)<>'00' "
        sql2=" and right(left(unit_code," & unit_len1 &"),2)<>'00' "
      else'ȫ��
        sql=" unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
        sql1=" and unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
        sql2=" and unit_code like '" & left(request("unit_code1"),unit_len1) & "%' and mid(unit_code,"&unit_len2+1&","&unit_len0-unit_len3&")<>'00'"
      end if
    else'�շ�վ
      sql=" unit_code='" & request("unit_code2") & "'"
      sql1=" and unit_code='" & request("unit_code2") & "'"
      sql2=" and unit_code='" & request("unit_code2") & "'"
    end if
    bbjzdwmc=""
    bbjzzg=""
    bbjzfh=""
    bbjzzb=""
    if request("unit_code2")="" then
      if right(left(request("unit_code1"),unit_len1),2)="00" then'ȫʡ
        rs.open "select * from unit where unit_code='" & request("unit_code1") & "'" ,conn,1,1
      else'ȫ��
        rs.open "select * from unit where unit_code='" & request("unit_code1") & "'" ,conn,1,1
      end if
    else'�շ�վ
        rs.open "select * from unit where unit_code='" & request("unit_code2") & "'" ,conn,1,1
    end if
    if rs.recordcount>0 then
      if not isnull(rs("bbjzdwmc")) then bbjzdwmc=rs("bbjzdwmc")
      if not isnull(rs("bbjzzg")) then bbjzzg=rs("bbjzzg")
      if not isnull(rs("bbjzfh")) then bbjzfh=rs("bbjzfh")
      if not isnull(rs("bbjzzb")) then bbjzzb=rs("bbjzzb")
    end if
    rs.close
    if trim(request("shj1")) <> "0" and trim(request("shj1"))<>"" then
      if sql<>"" then sql=sql & " and "
      if sql1<>"" then sql1=sql1 & " and "
      if sql2<>"" then sql2=sql2 & " and "
      sql=sql & " (shj1 like '" + trim(request("shj1")) + "%')"
      sql1=sql1 & " (shj1 like '" + trim(request("shj1")) + "%')"
      sql2=sql2 & " (shj1 like '" + trim(request("shj1")) + "%' and left(shj1,6)<='" + trim(request("shj1")) + trim(request("shj2")) +"')"
    end if
    'response.write sql
    rs.open "select * from edzdjb01 where " + sql, conn, 1, 1
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
            [<a href="searchzwqk.asp?mode=1">��������</a>]
          </td>
        </tr>
        <tr>
          <td align="center" colspan="2">
            ����ʡ��ͨ��·ͨ�з�����ͳ�Ʊ�
          </td>
        </tr>
        <tr>
          <td align="center" colspan="2">
            <%=bbjzdwmc & "-" & request("shj1")&"��" & request("shj2")& "��"%>
          </td>
        </tr>
        <tr>
          <td align="left">
            �౨��λ�����£�
          </td>
          <td align="right">
            �������ڣ�<%=year(now)%>��<%=month(now)%>��<%=day(now)%>��
          </td>
        </tr>
        <tr>
          <td colspan="2">
            <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
              <tr bgcolor=<%=skincolor()%>>
                <td align=center colspan="4">��Ŀ</td>
                <td width=140 align=center>���(��Ԫ)</td>
              </tr>
              <%
              byhj=0
              bnlj=0
              fl=true
              rs.close
              xxhs=0
              fdhs=0
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0101000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              xxhs=xxhs+rs.recordcount+1
              rs.close
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0102000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' and ajlb_code<>'0102050000' order by ajlb_sxh", conn, 1, 1
              xxhs=xxhs+rs.recordcount+1
              fdhs=fdhs+rs.recordcount+1
              rs.close
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len3 & ")='" & left("0102050000",ajlb_len3) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              xxhs=xxhs+rs.recordcount+1
              fdhs=fdhs+rs.recordcount+1
              rs.close
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0101000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              response.write "<tr>"
              response.write "<td rowspan=" & xxhs & " align=center>���й���ծ��</td>"
              if rs.recordcount>0 then
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x01.ajlbV) as expr1 from edzdjb01,edzdjb_x01,ajlb where edzdjb01.bh=edzdjb_x01.bh and edzdjb_x01.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<td align=center rowspan="&(rs.recordcount+1)&">������</td>"
                do while not rs.eof 
                  if not fl then
                    response.write "<tr>"
                    fl=true
                  end if
                  response.write "<td align=center colspan=2>" & rs("ajlb_name") & "</td>"
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
                  response.write "</tr>"
                  fl=false
                  rs.movenext
                loop
                rs1.close
                response.write "<tr>"
                response.write "<td align=center colspan=2>�ϼ�</td>"
                response.write "<td align=right>"&byhj&"</td>"
                response.write "</tr>"
              end if
              rs.close
              byhj=0
              bnlj=0
              fl=true
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0102000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' and ajlb_code<>'0102050000' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x01.ajlbV) as expr1 from edzdjb01,edzdjb_x01,ajlb where edzdjb01.bh=edzdjb_x01.bh and edzdjb_x01.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<tr>"
                response.write "<td align=center rowspan="&fdhs&">�������</td>"
                do while not rs.eof 
                  if not fl then
                    response.write "<tr>"
                    fl=true
                  end if
                  response.write "<td align=center colspan=2>" & rs("ajlb_name") & "</td>"
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
                  response.write "</tr>"
                  fl=false
                  rs.movenext
                loop
                rs1.close
                response.write "<tr>"
                response.write "<td align=center colspan=2>�ϼ�</td>"
                response.write "<td align=right>"&byhj&"</td>"
                response.write "</tr>"
              end if
              rs.close
              byhj=0
              bnlj=0
              fl=true
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len3 & ")='" & left("0102050000",ajlb_len3) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x01.ajlbV) as expr1 from edzdjb01,edzdjb_x01,ajlb where edzdjb01.bh=edzdjb_x01.bh and edzdjb_x01.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len3 & ")='" & left(rs("ajlb_code"),ajlb_len3) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len4) & ")='" & ajlb_str4 &"' and mid(ajlb.ajlb_code,"& (ajlb_len3+1) & "," & (ajlb_len4-ajlb_len3) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<tr>"
                response.write "<td align=center rowspan="&(rs.recordcount+1)&">����ɱ�</td>"
                do while not rs.eof 
                  if not fl then
                    response.write "<tr>"
                    fl=true
                  end if
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
                  response.write "</tr>"
                  fl=false
                  rs.movenext
                loop
                rs1.close
                response.write "<tr>"
                response.write "<td align=center>�ϼ�</td>"
                response.write "<td align=right>"&byhj&"</td>"
                response.write "</tr>"
              end if
              rs.close
              byhj=0
              bnlj=0
              fl=true
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0103000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x01.ajlbV) as expr1 from edzdjb01,edzdjb_x01,ajlb where edzdjb01.bh=edzdjb_x01.bh and edzdjb_x01.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<tr>"
                response.write "<td align=center rowspan="&(rs.recordcount+1)&">������Ŀ����</td>"
                do while not rs.eof 
                  if not fl then
                    response.write "<tr>"
                    fl=true
                  end if
                  response.write "<td align=center colspan=3>" & rs("ajlb_name") & "</td>"
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
                  response.write "</tr>"
                  fl=false
                  rs.movenext
                loop
                rs1.close
                response.write "<tr>"
                response.write "<td align=center colspan=3>�ϼ�</td>"
                response.write "<td align=right>"&byhj&"</td>"
                response.write "</tr>"
              end if
              rs.close
              byhj=0
              bnlj=0
              fl=true
              rs.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left("0104000000",ajlb_len2) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
              if rs.recordcount>0 then
                rs1.open "select ajlb.ajlb_code,sum(edzdjb_x01.ajlbV) as expr1 from edzdjb01,edzdjb_x01,ajlb where edzdjb01.bh=edzdjb_x01.bh and edzdjb_x01.ajlb_code=ajlb.ajlb_code and left(ajlb.ajlb_code," & ajlb_len2 & ")='" & left(rs("ajlb_code"),ajlb_len2) &"' and right(ajlb.ajlb_code,"& (ajlb_len0-ajlb_len3) & ")='" & ajlb_str3 &"' and mid(ajlb.ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' " & sql1 & " group by ajlb.ajlb_code", conn, 1, 1
                response.write "<tr>"
                response.write "<td align=center rowspan="&(rs.recordcount+1)&">��δ���</td>"
                do while not rs.eof 
                  if not fl then
                    response.write "<tr>"
                    fl=true
                  end if
                  response.write "<td align=center colspan=3>" & rs("ajlb_name") & "</td>"
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
                  response.write "</tr>"
                  fl=false
                  rs.movenext
                loop
                rs1.close
                response.write "<tr>"
                response.write "<td align=center colspan=3>�ϼ�</td>"
                response.write "<td align=right>"&byhj&"</td>"
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
                <td align=left>���ܣ�<%=bbjzzg%></td>
                <td align=center>���ˣ�<%=bbjzfh%></td>
                <td align=right>�Ʊ�<%=bbjzzb%></td>
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