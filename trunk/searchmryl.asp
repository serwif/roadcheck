<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>
<%
if session("visitor")=1 then
elseif session("username")=""  or (instr(session("power"),",1,")=0  and instr(session("power"),",2,")=0 and instr(session("power"),",3,")=0) then
  Response.Redirect("notlogin.asp")
end if
%>
<!--#include file="fcommon.asp"-->
<!--#include file="dtp.asp"-->
<%
dim conn, mode, username, rs, sql,rs1,rsMX,rs2,rs3, errmsg, founderror, s, t, i, fl, dq,odq, cpage,kpbm,st
dim unit_code,fph,tjlb_code,tjlb_str,tjlb_str_qt,afsj,ajjs,czshj,explain
dim DisSQL,sday,sday0,xjfas(),xjpas(),xjfps,zjfps,DisDate,days,ajjs_str,ajs
dim tjbb,fahj,pahj,fas,pas,fajs,pajs
dim tmpfbs,tmpsws,tmpclbfpl,tmpzs,sfhj

if not isempty(request("mode")) then
    mode = clng(request("mode"))
else
    mode=1
end if
'Response.Write visitor
if not isempty(request("username")) then
    username = request("username")
else
    username = ""
end if
if not isempty(request("DisDate")) then
    DisDate = request("DisDate")
else
    DisDate = 8
end if
if not isempty(request("sday")) then
    sday = request("sday")
else
    sday = date()
end if
if not isempty(request("DisSQL")) then
    DisSQL = request("DisSQL")
else
    DisSQL = ""
end if
if not isempty(request("fph")) then
    fph = request("fph")
else
    fph = ""
end if
if not isempty(request("unit_code")) then
    unit_code = request("unit_code")
else
    unit_code = ""
end if
if not isempty(request("tjlb_code")) then
    tjlb_code = request("tjlb_code")
else
    tjlb_code = ""
end if
if not isempty(request("afsj")) then
    afsj = request("afsj")
    if len(afsj)=10 and isdate(left(afsj,10))  then
      afsj=datetostr(left(afsj,10))
    end if
else
    afsj = ""
end if
if not isempty(request("ajjs")) then
    ajjs = request("ajjs")
else
    ajjs = ""
end if
if not isempty(request("tjbb")) then
    tjbb = request("tjbb")
else
    tjbb = ""
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
  <meta HTTP-EQUIV="Content-Type" content="text/html; charset=gb_2312">
  <meta http-equiv="Expires" CONTENT="0">
  <meta http-equiv="Cache-Control" CONTENT="no-cache">
  <meta http-equiv="Pragma" CONTENT="no-cache">
  <title>ÿ��һ��</title>
  <link rel="stylesheet" type="text/css" href="./main.css">
  </head>
<script language="javascript">
function trim(word) {
  while(word.length>0) {
    if(word.substring(0,1)==" ")
       word=word.substring(1,word.length)
    else
     if(word.substring(0,2)=="  ")
        word=word.substring(2,word.length)
     else
       break
}
   while(word.length>0) {
     if(word.substring(word.length-1,word.length)==".")
        word = word.substring(0,word.length-1)
     else
        if(word.substring(word.length-2,word.length)=="  ")
           word = word.substring(0,word.length-2)
        else
          break
 }
return word
}

function isnumber(word){
  var i=0;
  var result=true;

  for(i=0;i<word.length;i++){
    if(word.charAt(i)<'0'||word.charAt(i)>'9'){
	   result=false;
	   break;
  }
 }
return result;
} 

function check()
{
   //alert("searchmryl.asp?mode=1&DisDate=8&sday=" + document.all.afsj.value); 
   location.href="searchmryl.asp?mode=1&DisDate=8&sday=" + document.all.afsj.value+"&tjbb=" + document.all.tjbb.value; 
   return false;    
}

function check_form() 
{   
  if(trim(document.form1.ajsj.value)==""){
   alert("���������밸������!"); 
   return false; 
   }
  if(trim(document.form1.ajjs.value)==""){
   alert("�����밸������!"); 
   return false; 
   } 
return true;
} 

function hiddiv(blah) 
{ 
blah.style.display="none" 
} 
function showdiv(blah) 
{ 
blah.style.display="" 
blah.style.left=window.event.clientX+15 
blah.style.top=window.event.clientY 
} 

function showMsg(text) {
document.picform.message.value = text;
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

if mode=1 then
  '��ʾÿ��һ��
  'response.write Disdate
  if not isEmpty(request("page")) and isnumeric(request("page")) then
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
  set rs2=server.createobject("adodb.recordset")
  set rs3=server.createobject("adodb.recordset")
  sql=""
  if DisDate=4 then' ����
    sday0 = DateAdd("m", -1, sday)
    sday0 = Year(sday0) & "-" & Month(sday0) & "-1"
    sql=sql+" AND afsj>='" & datetostr(sday0) & "'"
    days=year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "�� "
    sday0 = DateAdd("d", -1, Year(sday) & "-" & Month(sday) & "-1")
    sql=sql+" AND afsj<='" & datetostr(sday0) & "'"
    days=days & "��" & year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "��"
  elseif DisDate=5 then' ����
    sday0=sday-weekday(sday,vbMonday)+2-7
    sql=sql+" AND afsj>='" & datetostr(sday0) & "' and afsj<='" & datetostr(sday0+6) &"'"
    days=year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "�� �� " & year(sday0+6) &"��" & month(sday0+6) &"��" & day(sday0+6) & "��"
  elseif DisDate=6 then' ����
    sday0 = Year(sday) & "-" & Month(sday) & "-1"
    sql=sql+" AND afsj>='" & datetostr(sday0) & "'"
    days=year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "�� "
    sday0=dateadd("m",1,sday)
    sday0 = DateAdd("d", -1, Year(sday0) & "-" & Month(sday0) & "-1")
    sql=sql+" AND afsj<='" & datetostr(sday0) & "'"
    days=days & "��" & year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "��"
  elseif DisDate=7 then' ����
    'response.write weekday(sday,vbMonday)'Weekday(date, [firstdayofweek]) Firstdayofweek ��ѡ��ָ��һ���ڵ�һ��ĳ��������δ��ָ�������� vbSunday Ϊȱʡֵ�� ���� firstdayofweek �����������趨ֵ�� ���� ֵ ���� bUseSystem 0 ʹ�� NLS API ���á� 
    'VbSunday 1 �����գ�ȱʡֵ�� 
    'vbMonday 2 ����һ 
    'vbTuesday 3 ���ڶ� 
    'vbWednesday 4 ������ 
    'vbThursday 5 ������ 
    'vbFriday 6 ������ 
    'vbSaturday 7 ������ 
    sday0=date()-weekday(sday,vbMonday)+2
    sql=sql+" AND afsj>='" & datetostr(sday0) & "' and afsj<='" & datetostr(sday0+6) &"'"
    days=year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "�� �� " & year(sday0+6) &"��" & month(sday0+6) &"��" & day(sday0+6) & "��"
  elseif DisDate=8 then'ĳ��
    sday0=sday
    sql=sql+" AND afsj='" & datetostr(sday0) & "'"
    days=year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "��"
  elseif DisDate=9 then'ȫ��
    sday0=year(sday) & "-01-01"
    'response.write sday0
    sql=sql+" AND afsj>='" & datetostr(sday0) & "'"
    days=year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "�� "
    sday0=year(sday) & "-12-31"
    'response.write sday0
    sql=sql+" AND afsj<='" & datetostr(sday0) & "'"
    days=days & "��" & year(sday0) &"��" & month(sday0) &"��" & day(sday0) & "��"
  end if
  dissql=sql
  'response.write sql
  %>
  <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td height="23" align="left" bgcolor="#FFFFFF" >
        <input type="text" name="afsj" size="10" maxlength="12" readonly  value='<%=sday%>' onchange="check()">
        <input type="hidden" name="tjbb" size="10" maxlength="12" readonly  value='<%=request("tjbb")%>'>
        <A onclick="show_cele_date(change2,'','',afsj)"><IMG align=top border=0 height=25 name=change2 src="images\calendar.gif" width=26></A>
        [<a href="searchmryl.asp?mode=1&DisDate=4&tjbb=<%=request("tjbb")%>">  ����  </a>]
        [<a href="searchmryl.asp?mode=1&DisDate=5&tjbb=<%=request("tjbb")%>">  ����  </a>]
        [<a href="searchmryl.asp?mode=1&DisDate=6&tjbb=<%=request("tjbb")%>">  ����  </a>]
        [<a href="searchmryl.asp?mode=1&DisDate=7&tjbb=<%=request("tjbb")%>">  ����  </a>]
        [<a href="searchmryl.asp?mode=1&DisDate=9&tjbb=<%=request("tjbb")%>">  ȫ��  </a>]
      </td>
    </tr>
  </tabel>
  <%rsMX.open "select * from tjlb where left(tjlb_code," & tjlb_len1 & ")='" & left(request("tjbb"),tjlb_len1) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len2) & ")='" & tjlb_str2 &"' and mid(tjlb_code,"& (tjlb_len1+1) & "," & (tjlb_len2-tjlb_len1) & ")<>'00' order by tjlb_sxh",conn,1,1
  rs1.open "select * from unit order by unit_sxh", conn, 1,1'?�о��Ƿ�Ҳ�а�������--û��
  if rsMX.recordcount<>0 then
    redim xjfas(rsMX.recordcount),xjpas(rsMX.recordcount)
    %>
    <br>
    <table width="95%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
      <tr bgcolor=<%=skincolor()%> height="28">
        <td align="center"><b><%
        sfhj="N"
        rs.open "select * from tjlb where tjlb_code='" + request("tjbb") + "'", conn, 1, 1
        if rs.recordcount>0 then
          days=days & rs("tjlb_name")
          if not isnull(rs("sfxsxj")) then
            sfhj=rs("sfxsxj")
          end if
        end if
        rs.close
        if sql<>"" then
          rs.open "select max(czshj) as expr1 from ajdjb where  afsj='" & datetostr(sday) &"'",conn,1,1
        else
          rs.open "select max(czshj) as expr1 from ajdjb where  afsj='" & datetostr(sday) & "'",conn,1,1
        end if
        if isnull(rs("expr1")) then

        elseif trim(rs("expr1"))="" then

        else
           days=days & " (������ʱ�䣺" & rs("expr1") & ")"
        end if
        rs.close
        Response.write days
        %></b></td>
        </tr>
        </table><br><!--
        <div align="center">
        &nbsp;&nbsp;<a href="clb_yqb5.asp">����ͳ��</a>
        </div>
        <br>--><!--������ֻ���˴�ӡͳ�Ʊ�������ͳ��,����Ҫ����ӡͳ�Ʊ����С��,���Ҫ����С���Ƿ���ʾ���ж�,���⻹Ҫ���������һ���ϼ���-->
        <table width="95%" border="0" cellspacing="1" cellpadding="5" bgcolor="#99CCFF" align="center">
          <tr bgcolor="#dbecec"> 
            <td colspan="1" rowspan="3" width=30 align="center">����</td><!--��-->
            <%rsMX.movefirst
            while not rsMX.eof'��ӡ��������,�������û��С��,������ռ��������;��������ռ1��(С������ʾ�ĸ���+1)*2��
              'rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and sfxsxj='Y' order by tjlb_sxh",conn,1,1
              rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and sfxsxj='Y' order by tjlb_sxh",conn,1,1
              if rs2.recordcount=0 then
                if not isnull(rsMX("sfxsxj")) then
                  if rsMX("sfxsxj")="Y" then%>
                    <td colspan="2" rowspan="2" align="center"><%=rsMX("tjlb_name")%></td>
                  <%end if%>
                <%end if%>
              <%else%>
                <%if not isnull(rsMX("sfxsxj")) then
                  if rsMX("sfxsxj")="Y" then%>
                    <td colspan="<%=(rs2.recordcount+1)*2%>" align="center"><%=rsMX("tjlb_name")%></td>
                  <%else%>
                    <td colspan="<%=(rs2.recordcount)*2%>" align="center"><%=rsMX("tjlb_name")%></td>
                  <%end if%>
                <%else%>
                  <td colspan="<%=(rs2.recordcount)*2%>" align="center"><%=rsMX("tjlb_name")%></td>
                <%end if%>
	      <%end if
              rs2.close
              rsMX.movenext
	    wend%>
            <%if sfhj="Y" then%>
              <td colspan="2" rowspan="2" align="center">�ϼ�</td>
            <%end if%>
          </tr>
          <tr bgcolor="#dbecec"> 
            <%rsMX.movefirst
            while not rsMX.eof'��ӡ����С��,�������û��С��,��ֻ��ӡС��(ռ1������);�����ȴ�ӡ��С��(ռ1������),�ٴ�ӡС��(ռ1������)
              'rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and sfxsxj='Y' order by tjlb_sxh",conn,1,1
              rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and sfxsxj='Y' order by tjlb_sxh",conn,1,1
              do while not rs2.eof%>
                <td colspan="2" align="center"><%=rs2("tjlb_name")%></td>
	      <%rs2.movenext
              loop
              if rs2.recordcount>0 then
                if not isnull(rsMX("sfxsxj")) then
                  if rsMX("sfxsxj")="Y" then%>
                    <td colspan="2" rowspan="1" align="center">С��</td>
                  <%end if
                end if
              end if
              rs2.close
              rsMX.movenext
	    wend%>
          </tr>
          <tr bgcolor="#dbecec"> 
            <%rsMX.movefirst
              while not rsMX.eof'��ӡ����С��,�������û��С��,��ֻ��ӡС��(ռ1������);�����ȴ�ӡ��С��ķ���,�ٴ�ӡС�Ƶķ���
              'rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and sfxsxj='Y' order by tjlb_sxh",conn,1,1
              rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and sfxsxj='Y' order by tjlb_sxh",conn,1,1
              do while not rs2.eof%>
                <td align="center">��</td>
                <td align="center">��</td>
	      <%rs2.movenext
              loop%>
              <%if rs2.recordcount=0 then%>
                <td align="center">��</td>
                <td align="center">��</td>
              <%else%>
                <%if isnull(rsMX("sfxsxj")) then%>
                  <td align="center">��</td>
                  <td align="center">��</td>
                <%else%>
                  <%if rsMX("sfxsxj")="N" then%>

                  <%else%>
                    <td align="center">��</td>
                    <td align="center">��</td>
                  <%end if%>
                <%end if%>
              <%end if%>
              <%rs2.close
              rsMX.movenext
	    wend%>
            <%if sfhj="Y" then%>
              <td align="center">��</td>
              <td align="center">��</td>
            <%end if%>
          </tr>
          <%while not rs1.EOF '����
            fahj=0
            pahj=0
            tjlb_str_qt=""
            %>
	    <tr bgcolor="#dbecec"> 
              <td colspan="1" rowspan="1" align="center">
                <%if right(rs1("unit_code"),2)="00" then%>
                  <!--<a href="searchmryl_bt.asp?DisDate=<%=DisDate%>&tjbb=<%=request("tjbb")%>&DisSQL=<%=DisSQL%>&unit_code=all&tjlb_code=all&btsm=ȫ�и��స����ͼ">ȫ��</a>-->
                  ȫ��
                <%else%>
                  <!--<a href="searchmryl_bt.asp?DisDate=<%=DisDate%>&tjbb=<%=request("tjbb")%>&DisSQL=<%=DisSQL%>&unit_code=<%=rs1("unit_code")%>&tjlb_code=all&btsm=<%=rs1("unit_name")%>���స����ͼ"><%=rs1("unit_name")%></a>-->
                  <%=rs1("unit_name")%>
                <%end if%>
              </td><!--��ӡ�������ƣ�����������Ϊ��Ӧ�Ķ�����Ŀ�ĸ���-->
	      <%rsmx.movefirst
              while not rsMX.eof
                fas=0
                pas=0
                fajs=""
                pajs=""
                if rsMX("bz")="-" then'��ͨ
                  'rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh",conn,1,1
                  rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh",conn,1,1
                  do while not rs2.eof
                    tjlb_str=""
                    rs3.open "select * from tjlb where left(tjlb_code," & tjlb_len3 & ")='" & left(rs2("tjlb_code"),tjlb_len3) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len4) & ")='" & tjlb_str4 &"' and mid(tjlb_code,"& (tjlb_len3+1) & "," & (tjlb_len4-tjlb_len3) & ")<>'00' order by tjlb_sxh",conn,1,1
                    if rs3.recordcount>0 then
                      do while not rs3.eof
                        if tjlb_str<>"" then tjlb_str=tjlb_str +" or "
                        if tjlb_str_qt<>"" then tjlb_str_qt=tjlb_str_qt+" or "
                        if rs3("gs_cc")=4 then'�������
                          tjlb_str=tjlb_str +" fxlb_code like '"+rs3("gs")+"%' "
                          tjlb_str_qt=tjlb_str_qt +" fxlb_code like '"+rs3("gs")+"%' "
                        else
                          tjlb_str=tjlb_str +" ajlb_code like '"+rs3("gs")+"%' "
                          tjlb_str_qt=tjlb_str_qt +" ajlb_code like '"+rs3("gs")+"%' "
                        end if
                        rs3.movenext
                      loop
                      if tjlb_str<>"" then tjlb_str=" and ("+tjlb_str+") "
                    else
                      if tjlb_str_qt<>"" then tjlb_str_qt=tjlb_str_qt+" or "
                      if rs2("gs_cc")=4 then'�������
                        tjlb_str=" and fxlb_code like '"+rs2("gs")+"%' "
                        tjlb_str_qt=tjlb_str_qt +" fxlb_code like '"+rs2("gs")+"%' "
                      else
                        tjlb_str=" and ajlb_code like '"+rs2("gs")+"%' "
                        tjlb_str_qt=tjlb_str_qt +" ajlb_code like '"+rs2("gs")+"%' "
                      end if
                    end if
                    rs3.close
                    '����
                    ajjs_str=""
                    ajs=0
                    if right(rs1("unit_code"),2)="00" then
                      rs.open "select ajjs from ajdjb where fph='FH' "+tjlb_str+" and unit_code like '" & left(rs1("unit_code"),4) & "%' " & sql &" order by bh",conn,1,1
                    else
                      rs.open "select ajjs from ajdjb where fph='FH' "+tjlb_str+" and unit_code='" & rs1("unit_code") & "' " & sql &" order by bh",conn,1,1
                    end if
                    'response.write "select ajjs from ajdjb where fph='FH' "+tjlb_str+" and unit_code='" & rs1("unit_code") & "' " & sql &" order by bh"
                    'response.write rs.recordcount
                    if rs.recordcount=0 then
     
                    else
                      do while not rs.eof
                        if ajjs_str<>"" then ajjs_str=ajjs_str+"<br>"
                        ajs=ajs+1
                        ajjs_str=ajjs_str+cstr(ajs)+":"+rs("ajjs")+";"
                        rs.movenext
                      loop
                    end if
                    rs.close
                    if rs2("sfxsxj")="Y" then'����ͳ�������ø�С���Ƿ���ʾ
                      if ajs=0 then'��ʾС�෢����
                        response.write("<td align=right>0</td>")
                      else
                        if right(rs1("unit_code"),2)="00" then
                          response.write("<td align=right>" & ajs &"</td>")
                        else%>
                          <td align=right><div id=f<%=rs1("unit_code")%><%=rs2("tjlb_code")%> style="position: absolute; width: 126; height: 27; background-color: orange; display: none; left: 11; top: 36"><%=ajjs_str%></div> <a href="searchmryl.asp?mode=2&tjbb=<%=request("tjbb")%>&DisDate=<%=DisDate%>&fph=FH&tjlb_code=<%=rs2("tjlb_code")%>&unit_code=<%=rs1("unit_code")%>&dissql=<%=dissql%>&tjlb_cc=3" onmouseout=hiddiv(f<%=rs1("unit_code")%><%=rs2("tjlb_code")%>); onmousemove=showdiv(f<%=rs1("unit_code")%><%=rs2("tjlb_code")%>);><%=ajs%></a></td>
                        <%end if
                      end if
                    end if
                    fas=fas+ajs
                    if ajjs_str<>"" then
                      if fajs<>"" then fajs=fajs+"<br>"
                      fajs=fajs+ajjs_str
                    end if
                    '�ư�
                    ajjs_str=""
                    ajs=0
                    if right(rs1("unit_code"),2)="00" then                  
                      rs.open "select ajjs from ajdjb where fph='PH' "+tjlb_str+" and unit_code like '" & left(rs1("unit_code"),4) & "%' " & sql &" order by bh",conn,1,1
                    else
                      rs.open "select ajjs from ajdjb where fph='PH' "+tjlb_str+" and unit_code='" & rs1("unit_code") & "' " & sql &" order by bh",conn,1,1
                    end if
                    if rs.recordcount=0 then
                      
                    else
                      do while not rs.eof
                        if ajjs_str<>"" then ajjs_str=ajjs_str+"<br>"
                        ajs=ajs+1
                        ajjs_str=ajjs_str+cstr(ajs)+":"+rs("ajjs")+";"
                        rs.movenext
                      loop
                    end if
                    rs.close
                    if rs2("sfxsxj")="Y" then'����ͳ�������ø�С���Ƿ���ʾ
                      if ajs=0 then'��ʾС���ư���
                        response.write("<td align=right>0</td>")
                      else
                        if right(rs1("unit_code"),2)="00" then
                          response.write("<td align=right>" & ajs &"</td>")
                        else%>
                          <td align=right><div id=p<%=rs1("unit_code")%><%=rs2("tjlb_code")%> style="position: absolute; width: 126; height: 27; background-color: orange; display: none; left: 11; top: 36"><%=ajjs_str%></div> <a href="searchmryl.asp?mode=2&tjbb=<%=request("tjbb")%>&DisDate=<%=DisDate%>&fph=FH&tjlb_code=<%=rs2("tjlb_code")%>&unit_code=<%=rs1("unit_code")%>&dissql=<%=dissql%>&tjlb_cc=3" onmouseout=hiddiv(p<%=rs1("unit_code")%><%=rs2("tjlb_code")%>); onmousemove=showdiv(p<%=rs1("unit_code")%><%=rs2("tjlb_code")%>);><%=ajs%></a></td>
                        <%end if
                      end if
                    end if
                    pas=pas+ajs
                    if ajjs_str<>"" then
                      if pajs<>"" then pajs=pajs+"<br>"
                      pajs=pajs+ajjs_str
                    end if
                    rs2.movenext
                  loop
                  rs2.close
                  fahj=fahj+fas
                  pahj=pahj+pas
                  if not isnull(rsMX("sfxsxj")) then
                    if rsMX("sfxsxj")="Y" then
                      if fas=0 then
                        response.write("<td align=right>0</td>")
                      else
                        if right(rs1("unit_code"),2)="00" then
                          response.write("<td align=right>" & fas &"</td>")
                        else%>
                          <td align=right><div id=f<%=rs1("unit_code")%><%=rsMX("tjlb_code")%> style="position: absolute; width: 126; height: 27; background-color: orange; display: none; left: 11; top: 36"><%=fajs%></div> <a href="searchmryl.asp?mode=2&tjbb=<%=request("tjbb")%>&DisDate=<%=DisDate%>&fph=FH&tjlb_code=<%=rsMX("tjlb_code")%>&unit_code=<%=rs1("unit_code")%>&dissql=<%=dissql%>&tjlb_cc=2" onmouseout=hiddiv(f<%=rs1("unit_code")%><%=rsMX("tjlb_code")%>); onmousemove=showdiv(f<%=rs1("unit_code")%><%=rsMX("tjlb_code")%>);><%=fas%></a></td>
                        <%end if
                      end if
                    end if
                  end if
                  if not isnull(rsMX("sfxsxj")) then
                    if rsMX("sfxsxj")="Y" then
                      if pas=0 then
                        response.write("<td align=right>0</td>")
                      else
                        if right(rs1("unit_code"),2)="00" then
                          response.write("<td align=right>" & pas &"</td>")
                        else%>
                          <td align=right><div id=p<%=rs1("unit_code")%><%=rsMX("tjlb_code")%> style="position: absolute; width: 126; height: 27; background-color: orange; display: none; left: 11; top: 36"><%=pajs%></div> <a href="searchmryl.asp?mode=2&tjbb=<%=request("tjbb")%>&DisDate=<%=DisDate%>&fph=PH&tjlb_code=<%=rsMX("tjlb_code")%>&unit_code=<%=rs1("unit_code")%>&dissql=<%=dissql%>&tjlb_cc=2" onmouseout=hiddiv(p<%=rs1("unit_code")%><%=rsMX("tjlb_code")%>); onmousemove=showdiv(p<%=rs1("unit_code")%><%=rsMX("tjlb_code")%>);><%=pas%></a></td>
                        <%end if
                      end if
                    end if
                  end if
                elseif rsMX("bz")="QT" then'����
                  tjlb_str=""
                  if tjlb_str_qt<>"" then tjlb_str=" and not (" +tjlb_str_qt +") "
                  'response.write tjlb_str
                  '����
                  ajjs_str=""
                  ajs=0
                  if right(rs1("unit_code"),2)="00" then
                    rs.open "select ajjs from ajdjb where fph='FH' "+tjlb_str+" and unit_code like '" & left(rs1("unit_code"),4) & "%' " & sql &" order by bh",conn,1,1
                  else
                    rs.open "select ajjs from ajdjb where fph='FH' "+tjlb_str+" and unit_code='" & rs1("unit_code") & "' " & sql &" order by bh",conn,1,1
                  end if
                  'response.write "select ajjs from ajdjb where fph='FH' "+tjlb_str+" and unit_code='" & rs1("unit_code") & "' " & sql &" order by bh"
                  'response.write rs.recordcount
                  if rs.recordcount=0 then
                  else
                    do while not rs.eof
                      if ajjs_str<>"" then ajjs_str=ajjs_str+"<br>"
                      ajs=ajs+1
                      ajjs_str=ajjs_str+cstr(ajs)+":"+rs("ajjs")+";"
                      rs.movenext
                    loop
                  end if
                  rs.close
                  fas=fas+ajs
                  if ajs=0 then'��ʾС�෢����
                    response.write("<td align=right>0</td>")
                  else
                    if right(rs1("unit_code"),2)="00" then
                      response.write("<td align=right>" & ajs &"</td>")
                    else%>
                      <td align=right><div id=f<%=rs1("unit_code")%>qt style="position: absolute; width: 126; height: 27; background-color: orange; display: none; left: 11; top: 36"><%=ajjs_str%></div> <a href="searchmryl.asp?mode=2&tjbb=<%=request("tjbb")%>&DisDate=<%=DisDate%>&fph=FH&tjlb_code=QT&unit_code=<%=rs1("unit_code")%>&dissql=<%=dissql%>&tjlb_cc=3" onmouseout=hiddiv(f<%=rs1("unit_code")%>qt); onmousemove=showdiv(f<%=rs1("unit_code")%>qt);><%=ajs%></a></td>
                    <%end if
                  end if
                  '�ư�
                  ajjs_str=""
                  ajs=0
                  if right(rs1("unit_code"),2)="00" then                  
                    rs.open "select ajjs from ajdjb where fph='PH' "+tjlb_str+" and unit_code like '" & left(rs1("unit_code"),4) & "%' " & sql &" order by bh",conn,1,1
                  else
                    rs.open "select ajjs from ajdjb where fph='PH' "+tjlb_str+" and unit_code='" & rs1("unit_code") & "' " & sql &" order by bh",conn,1,1
                  end if
                  if rs.recordcount=0 then
                    
                  else
                    do while not rs.eof
                      if ajjs_str<>"" then ajjs_str=ajjs_str+"<br>"
                      ajs=ajs+1
                      ajjs_str=ajjs_str+cstr(ajs)+":"+rs("ajjs")+";"
                      rs.movenext
                    loop
                  end if
                  rs.close
                  pas=pas+ajs
                  if ajs=0 then'��ʾС���ư���
                    response.write("<td align=right>0</td>")
                  else
                    if right(rs1("unit_code"),2)="00" then
                      response.write("<td align=right>" & ajs &"</td>")
                    else%>
                      <td align=right><div id=p<%=rs1("unit_code")%>qt style="position: absolute; width: 126; height: 27; background-color: orange; display: none; left: 11; top: 36"><%=ajjs_str%></div> <a href="searchmryl.asp?mode=2&tjbb=<%=request("tjbb")%>&DisDate=<%=DisDate%>&fph=FH&tjlb_code=QT&unit_code=<%=rs1("unit_code")%>&dissql=<%=dissql%>&tjlb_cc=3" onmouseout=hiddiv(p<%=rs1("unit_code")%>qt); onmousemove=showdiv(p<%=rs1("unit_code")%>qt);><%=ajs%></a></td>
                    <%end if
                  end if
                  fahj=fahj+fas
                  pahj=pahj+pas
                end if
                rsMX.movenext
              wend%>
              <%if sfhj="Y" then%>
                <td align="center"><%=fahj%></td>
                <td align="center"><%=pahj%></td>
              <%end if%>
	    </tr> 
	    <%rs1.movenext
	  wend%>
        </table>
      </table>
      <br>
      <%
      showctail()  
    else
      showchead()
      %>
      </p>
      </div>
      <p><br>
      </p><table width="95%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
      <tr bgcolor=<%=skincolor()%> height="28">
        <td align="center"><b>
         ȫ�и��స���������
        </b></td>
      </tr>
      <tr><td align=center><br>
      ����֪ͨϵͳ����Ա����ͳ�Ʊ����ֵ䣡
      <br><br></td></tr>
      </table>
      <br>
      <%
    end if
    set rs1=nothing
    rsMX.close
    set rsMX=nothing     
    set rs=nothing
    closedb()
    showctail()
elseif mode=2 then
  '��ʾĳ��ĳ�స������/�ư��İ�����¼
  'response.write(sday)
  if not isEmpty(request("page")) and isnumeric(request("page")) then
    cpage = clng(request("page"))
  else
    cpage = 1
  end if
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  set rsMX=server.createobject("adodb.recordset")
  set rs1=server.createobject("adodb.recordset")
  set rs2=server.createobject("adodb.recordset")
  set rs3=server.createobject("adodb.recordset")
  'response.write request("tjlb_code")
  if request("tjlb_code")="QT" then'����
    tjlb_str_qt=""
    rsMX.open "select * from tjlb where left(tjlb_code," & tjlb_len1 & ")='" & left(request("tjbb"),tjlb_len1) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len2) & ")='" & tjlb_str2 &"' and mid(tjlb_code,"& (tjlb_len1+1) & "," & (tjlb_len2-tjlb_len1) & ")<>'00' order by tjlb_sxh",conn,1,1
    do while not rsMX.eof
      tjlb_str=""
      rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(rsMX("tjlb_code"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh",conn,1,1
      'response.write rs2.recordcount
      do while not rs2.eof
        rs3.open "select * from tjlb where left(tjlb_code," & tjlb_len3 & ")='" & left(rs2("tjlb_code"),tjlb_len3) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len4) & ")='" & tjlb_str4 &"' and mid(tjlb_code,"& (tjlb_len3+1) & "," & (tjlb_len4-tjlb_len3) & ")<>'00' order by tjlb_sxh",conn,1,1
        if rs3.recordcount>0 then
          do while not rs3.eof
            if tjlb_str<>"" then tjlb_str=tjlb_str +" or "
            if tjlb_str_qt<>"" then tjlb_str_qt=tjlb_str_qt +" or "
            if rs3("gs_cc")=4 then'�������
              tjlb_str=tjlb_str +" fxlb_code like '"+rs3("gs")+"%' "
              tjlb_str_qt=tjlb_str_qt +" fxlb_code like '"+rs3("gs")+"%' "
            else
              tjlb_str=tjlb_str +" ajlb_code like '"+rs3("gs")+"%' "
              tjlb_str_qt=tjlb_str_qt +" ajlb_code like '"+rs3("gs")+"%' "
            end if
            rs3.movenext
          loop
        else
          if tjlb_str<>"" then tjlb_str=tjlb_str +" or "
          if tjlb_str_qt<>"" then tjlb_str_qt=tjlb_str_qt +" or "
          if rs2("gs_cc")=4 then'�������
            tjlb_str=tjlb_str+" fxlb_code like '"+rs2("gs")+"%' "
            tjlb_str_qt=tjlb_str_qt +" fxlb_code like '"+rs2("gs")+"%' "
          else
            tjlb_str=tjlb_str+" ajlb_code like '"+rs2("gs")+"%' "
            tjlb_str_qt=tjlb_str_qt +" ajlb_code like '"+rs2("gs")+"%' "
          end if
        end if
        rs3.close
        rs2.movenext
      loop
      rs2.close
      if tjlb_str<>"" then tjlb_str=" and ("+tjlb_str+") "
      rsMX.movenext
    loop
    rsMX.close
    tjlb_str=""
    if tjlb_str_qt<>"" then tjlb_str=" and not (" +tjlb_str_qt +") "
  else'��ͨ
    if request("tjlb_cc")=2 then
      tjlb_str=""
      rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(request("tjlb_code"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh",conn,1,1
      'response.write rs2.recordcount
      do while not rs2.eof
        rs3.open "select * from tjlb where left(tjlb_code," & tjlb_len3 & ")='" & left(rs2("tjlb_code"),tjlb_len3) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len4) & ")='" & tjlb_str4 &"' and mid(tjlb_code,"& (tjlb_len3+1) & "," & (tjlb_len4-tjlb_len3) & ")<>'00' order by tjlb_sxh",conn,1,1
        if rs3.recordcount>0 then
          do while not rs3.eof
            if tjlb_str<>"" then tjlb_str=tjlb_str +" or "
            if rs3("gs_cc")=4 then'�������
              tjlb_str=tjlb_str +" fxlb_code like '"+rs3("gs")+"%' "
            else
              tjlb_str=tjlb_str +" ajlb_code like '"+rs3("gs")+"%' "
            end if
            rs3.movenext
          loop
        else
          if tjlb_str<>"" then tjlb_str=tjlb_str +" or "
          if rs2("gs_cc")=4 then'�������
            tjlb_str=tjlb_str+" fxlb_code like '"+rs2("gs")+"%' "
          else
            tjlb_str=tjlb_str+" ajlb_code like '"+rs2("gs")+"%' "
          end if
        end if
        rs3.close
        rs2.movenext
      loop
      rs2.close
      if tjlb_str<>"" then tjlb_str=" and ("+tjlb_str+") "
    else
      tjlb_str=""
      'rs2.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(request("tjlb_code"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh",conn,1,1
      rs2.open "select * from tjlb where tjlb_code='" & request("tjlb_code") &"' order by tjlb_sxh",conn,1,1
      'response.write rs2.recordcount
      do while not rs2.eof
        rs3.open "select * from tjlb where left(tjlb_code," & tjlb_len3 & ")='" & left(rs2("tjlb_code"),tjlb_len3) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len4) & ")='" & tjlb_str4 &"' and mid(tjlb_code,"& (tjlb_len3+1) & "," & (tjlb_len4-tjlb_len3) & ")<>'00' order by tjlb_sxh",conn,1,1
        if rs3.recordcount>0 then
          do while not rs3.eof
            if tjlb_str<>"" then tjlb_str=tjlb_str +" or "
            if rs3("gs_cc")=4 then'�������
              tjlb_str=tjlb_str +" fxlb_code like '"+rs3("gs")+"%' "
            else
              tjlb_str=tjlb_str +" ajlb_code like '"+rs3("gs")+"%' "
            end if
            rs3.movenext
          loop
        else
          if tjlb_str<>"" then tjlb_str=tjlb_str +" or "
          if rs2("gs_cc")=4 then'�������
            tjlb_str=tjlb_str+" fxlb_code like '"+rs2("gs")+"%' "
          else
            tjlb_str=tjlb_str+" ajlb_code like '"+rs2("gs")+"%' "
          end if
        end if
        rs3.close
        rs2.movenext
      loop
      rs2.close
      if tjlb_str<>"" then tjlb_str=" and ("+tjlb_str+") "
    end if
  end if
  sql="select * from ajdjb where fph='"+fph+"' and unit_code='" + unit_code +"'" & dissql & tjlb_str
  'response.write(sql)
  rs.open sql, conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage>rs.pagecount then cpage=1
    rs.AbsolutePage = cpage
    %>
      <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
	</tr>
        <tr>
          <td valign="bottom">��<%=cstr(cpage)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>��������¼</td>
          <td align="right">
          [<a href="searchmryl.asp?mode=1&tjbb=<%=request("tjbb")%>&DisDate=<%=disdate%>">����</a>]
          <%if cpage <> 1 then%>
            [<a href="searchmryl.asp?mode=1&tjbb=<%=request("tjbb")%>&DisDate=<%=disdate%>&fph=<%=fph%>&tjlb_code=<%=tjlb_code%>&unit_code=<%=unit_code%>&page=<%=cstr(cpage-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage <> rs.PageCount then%>
            [<a href="searchmryl.asp?mode=1&tjbb=<%=request("tjbb")%>&DisDate=<%=disdate%>&page=<%=cstr(cpage+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage then%>
                <option selected value="searchmryl.asp?mode=1&tjbb=<%=request("tjbb")%>&DisDate=<%=disdate%>&page=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="seachmryl.asp?mode=1&DisDate=<%=disdate%>&page=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>�������</td>
              <td width=10% align=center>����ʱ��</td>
              <td width=30% align=center>��������</td>
              <!--<td width=15% align=center>����</td>-->
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
              <td align=center>
                <%if rs("fph")="FH" then 
                  response.write("������")
                elseif rs("fph")="PH" then
                  response.write("�ư���")
                end if
                rs1.open "select ajlb_name from ajlb where ajlb_code='"+rs("ajlb_code")+"'",conn,1,1
		if not rs1.eof then
		  response.write(rs1("ajlb_name"))
		end if
		rs1.close%>
              </td>
              <td align=center><%=todatestr(left(rs("afsj"),8))%></td> 
	      <td align=left>
	        <%if len(rs("ajjs"))>50 then 
		  Response.Write(left(rs("ajjs"),50)+"...")
                else
		  Response.Write(rs("ajjs"))
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
  <%else%>
    <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          [<a href="searchmryl.asp?mode=1&tjbb=<%=request("tjbb")%>&DisDate=<%=disdate%>">����</a>]
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">û�м�¼</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()
elseif mode=3 then
'��ʾ����

end if
%>    