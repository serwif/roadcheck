<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if session("username")="" or instr(session("power"),",0,")=0 then
    Response.Redirect "notlogin.asp"
end if
%>

<!--#include file="fcommon.asp"-->

<%
dim conn, mode, username, rs,rs1,rsMX, sql, errmsg, founderror, s, t, i, fl, dq,odq,dq0, dq1,cpage1,cpage2,cpage3,cpage4,kpbm,st,dwxh,sfzs,dqcode1,dqcode2,dqcode3,dqcode4,dqname1,dqname2,dqname3,dqname4,ajlb_code,fxlb_code,gs_cc,dqbz,dqgs,sfxsxj,str
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
if not isempty(request("dqcode1")) then
    dqcode1 = request("dqcode1")
else
    dqcode1 = ""
end if
if not isempty(request("dqcode2")) then
    dqcode2 = request("dqcode2")
else
    dqcode2 = ""
end if
if not isempty(request("dqcode3")) then
    dqcode3 = request("dqcode3")
else
    dqcode3 = ""
end if
if not isempty(request("dqcode4")) then
    dqcode4 = request("dqcode4")
else
    dqcode4 = ""
end if
if not isempty(request("dqname1")) then
    dqname1 = request("dqname1")
else
    dqname1 = ""
end if
if not isempty(request("dqname2")) then
    dqname2 = request("dqname2")
else
    dqname2 = ""
end if
if not isempty(request("dqname3")) then
    dqname3 = request("dqname3")
else
    dqname3 = ""
end if
if not isempty(request("dqname4")) then
    dqname4 = request("dqname4")
else
    dqname4 = ""
end if
if not isempty(request("ajlb_code")) then
    ajlb_code = request("ajlb_code")
else
    ajlb_code = ""
end if
if not isempty(request("fxlb_code")) then
    fxlb_code = request("fxlb_code")
else
    fxlb_code = ""
end if
if not isempty(request("gs_cc")) then
    gs_cc = request("gs_cc")
else
    gs_cc = ""
end if
  if not isEmpty(request("page1")) and isnumeric(request("page1")) then
    cpage1 = clng(request("page1"))
  else
    cpage1 = 1
  end if
  if not isEmpty(request("page2")) and isnumeric(request("page2")) then
    cpage2 = clng(request("page2"))
  else
    cpage2 = 1
  end if
  if not isEmpty(request("page3")) and isnumeric(request("page3")) then
    cpage3 = clng(request("page3"))
  else
    cpage3 = 1
  end if
  if not isEmpty(request("page4")) and isnumeric(request("page4")) then
    cpage4 = clng(request("page4"))
  else
    cpage4 = 1
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
  <title>ͳ�Ʊ������</title>
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
   var ajlb_code;
   if(j==1)
     ajlb_code=document.input1.ajlb_code1.options[document.input1.ajlb_code1.selectedIndex].value;
   else
     {if(j==2)
        ajlb_code=document.input1.ajlb_code2.options[document.input1.ajlb_code2.selectedIndex].value; 
      else
        {if(j==3)
           ajlb_code=document.input1.ajlb_code3.options[document.input1.ajlb_code3.selectedIndex].value; 
        } 
     }
   if(i==102)
     location.href="marea-tjbb.asp?mode=102&page1="+document.input1.page1.value+"&dqcode1="+document.input1.dqcode1.value+"&dqname1="+document.input1.dqname1.value+"&ajlb_code="+ajlb_code+"&gs_cc="+document.input1.ajlb_code0.options[document.input1.ajlb_code0.selectedIndex].value;
   else
     {if(i==103)
        location.href="marea-tjbb.asp?mode=103&page1="+document.input1.page1.value+"&dqcode1="+document.input1.dqcode1.value+"&dqname1="+document.input1.dqname1.value+"&ajlb_code="+ajlb_code+"&gs_cc="+document.input1.ajlb_code0.options[document.input1.ajlb_code0.selectedIndex].value+"&odq="+document.input1.odq.value;
      else
        {if(i==202)
           location.href="marea-tjbb.asp?mode=202&page1="+document.input1.page1.value+"&dqcode1="+document.input1.dqcode1.value+"&dqname1="+document.input1.dqname1.value+"&page2="+document.input1.page2.value+"&dqcode2="+document.input1.dqcode2.value+"&dqname2="+document.input1.dqname2.value+"&ajlb_code="+ajlb_code+"&gs_cc="+document.input1.ajlb_code0.options[document.input1.ajlb_code0.selectedIndex].value+"&gs_cc2="+document.input1.gs_cc2.value+"&gs2="+document.input1.gs2.value;
         else
           {if(i==203)
              location.href="marea-tjbb.asp?mode=203&page1="+document.input1.page1.value+"&dqcode1="+document.input1.dqcode1.value+"&dqname1="+document.input1.dqname1.value+"&page2="+document.input1.page2.value+"&dqcode2="+document.input1.dqcode2.value+"&dqname2="+document.input1.dqname2.value+"&ajlb_code="+ajlb_code+"&gs_cc="+document.input1.ajlb_code0.options[document.input1.ajlb_code0.selectedIndex].value+"&gs_cc2="+document.input1.gs_cc2.value+"&gs2="+document.input1.gs2.value+"&odq="+document.input1.odq.value;
            else
              {if(i==302)
                 location.href="marea-tjbb.asp?mode=302&page1="+document.input1.page1.value+"&dqcode1="+document.input1.dqcode1.value+"&dqname1="+document.input1.dqname1.value+"&page2="+document.input1.page2.value+"&dqcode2="+document.input1.dqcode2.value+"&dqname2="+document.input1.dqname2.value+"&page3="+document.input1.page3.value+"&dqcode3="+document.input1.dqcode3.value+"&dqname3="+document.input1.dqname3.value+"&ajlb_code="+ajlb_code+"&gs_cc="+document.input1.ajlb_code0.options[document.input1.ajlb_code0.selectedIndex].value+"&gs_cc2="+document.input1.gs_cc2.value+"&gs2="+document.input1.gs2.value+"&gs_cc3="+document.input1.gs_cc3.value+"&gs3="+document.input1.gs3.value;
               else
                 location.href="marea-tjbb.asp?mode=303&page1="+document.input1.page1.value+"&dqcode1="+document.input1.dqcode1.value+"&dqname1="+document.input1.dqname1.value+"&page2="+document.input1.page2.value+"&dqcode2="+document.input1.dqcode2.value+"&dqname2="+document.input1.dqname2.value+"&page3="+document.input1.page3.value+"&dqcode3="+document.input1.dqcode3.value+"&dqname3="+document.input1.dqname3.value+"&ajlb_code="+ajlb_code+"&gs_cc="+document.input1.ajlb_code0.options[document.input1.ajlb_code0.selectedIndex].value+"&gs_cc2="+document.input1.gs_cc2.value+"&gs2="+document.input1.gs2.value+"&gs_cc3="+document.input1.gs_cc3.value+"&gs3="+document.input1.gs3.value+"&odq="+document.input1.odq.value;
              }
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
      <%if mode<100 then %><!--��������-->
        <b>����ͳ�Ʊ�������</b>
      <%elseif mode>100 and mode<200 then %><!--��������-->
        <b>����[<%=request("dqname1")%>]-����ͳ�Ʊ�������</b>
      <%elseif mode>200 and mode<300 then %><!--С������-->
        <b>����[<%=request("dqname1")%>]-����[<%=request("dqname2")%>]-С��ͳ�Ʊ�������</b>
      <%elseif mode>300 and mode<400 then %><!--С������-->
        <b>����[<%=request("dqname1")%>]-����[<%=request("dqname2")%>]-С��[<%=request("dqname3")%>]-С1��ͳ�Ʊ�������</b>
      <%end if%>
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
  if mode = 2 then
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb order by ajlb_sxh", conn, 1, 1%>
    <form method="POST" action="marea-tjbb.asp?mode=2&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from tjlb where tjlb_code='" + request("odq") + "'", conn, 1, 1
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb order by ajlb_sxh", conn, 1, 1
    %>
    <form method="POST" action="marea-tjbb.asp?mode=3&page1=<%=cpage1%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-tjbb.asp?mode=1&page1=<%=cpage1%>">����</a>]
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
            <td colspan="3">��������࣬Ȼ������OK��</td>
          <%else%>
            <td colspan="3">��༭���࣬Ȼ������OK��</td>
          <%end if%>
        <%end if%>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=dq size=15 maxlength=<%=tjlb_len1%> class="smallInput" value='<%=request("dq")%>'>
            <%else%>
              <input name=dq size=15 maxlength=<%=tjlb_len1%> class="smallInput" value='<%=trim(left(rs("tjlb_code"),tjlb_len1))%>'>
            <%end if%>
            <font color=red>(*)</font>(��������ǰ<%=tjlb_len1%>λ,��<%=tjlb_len0-tjlb_len1%>λȫΪ0)
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input name=dq0 size=15 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=15 maxlength=30 class="smallInput" value='<%=trim(rs("tjlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�Ƿ���ʾ�ϼ�&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=2 then%>
              <input type="checkbox" name="sfxsxj" value='yes'>
            <%else%>
              <%if rs("sfxsxj")="Y" then%>
                <input type="checkbox" name="sfxsxj" value='yes' checked>
              <%else%>
                <input type="checkbox" name="sfxsxj" value='yes'>
              <%end if%>
            <%end if%>
          </td>
        </tr>
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
  end if
  showctail
end sub

sub ShowInputForm101(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 102 then
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh", conn, 1, 1
    %>
    <form method="POST" action="marea-tjbb.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh", conn, 1, 1
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from tjlb where tjlb_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-tjbb.asp?mode=103&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>">����</a>]
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
              <% if mode = 102 then%>
                ���������࣬Ȼ������OK��
              <%else%>
                ��༭���࣬Ȼ������OK��
                <input name="odq" type="hidden" value="<%=request("odq")%>">
              <%end if%>
            <%end if%>
            <input name="page1" type="hidden" value="<%=cpage1%>">
            <input name="dqcode1" type="hidden" value="<%=request("dqcode1")%>">
            <input name="dqname1" type="hidden" value="<%=request("dqname1")%>">
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=dq size=15 maxlength=<%=tjlb_len3-tjlb_len2%> class="smallInput" value='<%=request("dq")%>'>(ǰ<%=tjlb_len1%>λΪ<%=left(request("dqcode1"),tjlb_len1)%>,�����<%=tjlb_len3-tjlb_len2%>λ)
            <%else%>
              <input name=dq size=15 maxlength=<%=tjlb_len3-tjlb_len2%> class="smallInput" value='<%=trim(mid(rs("tjlb_code"),tjlb_len1+1,tjlb_len2-tjlb_len1))%>'>(ǰ<%=tjlb_len1%>λΪ<%=left(request("dqcode1"),tjlb_len1)%>,�����<%=tjlb_len3-tjlb_len2%>λ)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("tjlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�Ƿ���ʾС��&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <input type="checkbox" name="sfxsxj" value='yes'>
            <%else%>
              <%if rs("sfxsxj")="Y" then%>
                <input type="checkbox" name="sfxsxj" value='yes' checked>
              <%else%>
                <input type="checkbox" name="sfxsxj" value='yes'>
              <%end if%>
            <%end if%>
          </td>
        </tr>  
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��Ŀ����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=102 then%>
              <select name="FRMbz" OnClick="javascript:if(document.input1.FRMbz.value*1==0){test1.style.display=''}else{test1.style.display='none'}">
                <option value="0">��ͨ</option>
                <option value="1">����</option>
              </select>
            <%else%>
              <select name="FRMbz" OnClick="javascript:if(document.input1.FRMbz.value*1==0){test1.style.display=''}else{test1.style.display='none'}">
                <option value="0"<%if rs("bz")="-" then%>selected<%end if%>>��ͨ</option>
                <option value="1"<%if rs("bz")="QT" then%>selected<%end if%>>����</option>
              </select>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��ʽ&nbsp;</td>
          <td align=left colspan=2>
             <%if mode=102 then%>
               <DIV ID="test1" Style="position:relative; display:'';"> 
               ������������:
               <br>���
               <select name="ajlb_code0" style="HEIGHT:17px;WIDTH:119px">
                 <option value="1"<%if gs_cc="1" then %> selected <% end if %>>����</option>
                 <option value="2"<%if gs_cc="2" then %> selected <% end if %>>����</option>
                 <option value="3"<%if gs_cc="3" then %> selected <% end if %>>С��</option>
                 <option value="4"<%if gs_cc="4" then %> selected <% end if %>>�������</option>
               </select>
               <br>����
               <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(102,1)">
               <%while not rs1.EOF 
                 if trim(ajlb_code)="" then ajlb_code=trim(rs1("ajlb_code"))%>
                 <option value="<%=trim(rs1("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len1)=left(rs1("ajlb_code"),ajlb_len1) then %> selected <% end if %>><%=trim(rs1("ajlb_name"))%></option>
                 <%rs1.MoveNext 
               WEND%>
               </select>
               <br>����
               <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(102,2)">
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(ajlb_code,ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF
                 if mid(ajlb_code,ajlb_len1+1,ajlb_len2-ajlb_len1)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len2)=left(rsMX("ajlb_code"),ajlb_len2) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>С��
               <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(102,3)">
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(ajlb_code,ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF
                 if mid(ajlb_code,ajlb_len2+1,ajlb_len3-ajlb_len2)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len3)=left(rsMX("ajlb_code"),ajlb_len3) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>�������
               <select name="ajlb_code4" style="HEIGHT:17px;WIDTH:119px" >
               <option value="" <%if fxlb_code="" then %> selected <% end if %>></option>
               <%rsMX.open "select * from fxlb where left(fxlb_code," & ajlb_len3 & ")='" & left(ajlb_code,ajlb_len3) &"' order by fxlb_sxh", conn, 1, 1
               while not rsMX.EOF%>
                 <option value="<%=trim(rsMX("fxlb_code"))%>"<%if left(fxlb_code,fxlb_len1)=left(rsMX("fxlb_code"),fxlb_len1-ajlb_len3) then %> selected <% end if %>><%=trim(rsMX("fxlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               </DIV>
             <%else%>
               <%if rs("bz")="-" or rs("bz")="QNTQ" then%>
                 <DIV ID="test1" Style="position:relative; display:'';"> 
               <%else%>
                 <DIV ID="test1" Style="position:relative; display:'none';"> 
               <%end if%>
               ������������
               <br>���
               <select name="ajlb_code0" style="HEIGHT:17px;WIDTH:119px">
                 <option value="1"<%if gs_cc="1" then %> selected <% end if %>>����</option>
                 <option value="2"<%if gs_cc="2" then %> selected <% end if %>>����</option>
                 <option value="3"<%if gs_cc="3" then %> selected <% end if %>>С��</option>
                 <option value="4"<%if gs_cc="4" then %> selected <% end if %>>�������</option>
               </select>
               <br>����
               <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(103,1)">
               <%while not rs1.EOF
                 if trim(ajlb_code)="" then ajlb_code=trim(rs1("ajlb_code"))%> 
                 <option value="<%=trim(rs1("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len1)=left(rs1("ajlb_code"),ajlb_len1) then %> selected <% end if %>><%=trim(rs1("ajlb_name"))%></option>
                 <%rs1.MoveNext 
               WEND%>
               </select>
               <br>����
               <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(103,2)">
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(ajlb_code,ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF
                 if mid(ajlb_code,ajlb_len1+1,ajlb_len2-ajlb_len1)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len2)=left(rsMX("ajlb_code"),ajlb_len2) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>С��
               <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(103,3)">
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(ajlb_code,ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF
                 if mid(ajlb_code,ajlb_len2+1,ajlb_len3-ajlb_len2)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len3)=left(rsMX("ajlb_code"),ajlb_len3) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>�������
               <select name="ajlb_code4" style="HEIGHT:17px;WIDTH:119px" >
               <option value="" <%if fxlb_code="" then %> selected <% end if %>></option>
               <%rsMX.open "select * from fxlb where left(fxlb_code," & ajlb_len3 & ")='" & left(ajlb_code,ajlb_len3) &"' order by fxlb_sxh", conn, 1, 1
               while not rsMX.EOF%>
                 <option value="<%=trim(rsMX("fxlb_code"))%>"<%if left(fxlb_code,fxlb_len1)=left(rsMX("fxlb_code"),fxlb_len1-ajlb_len3) then %> selected <% end if %>><%=trim(rsMX("fxlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               </DIV>
             <%end if%>
           </td>
        </tr>
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
  if mode = 103 then
    rs1.close
    set rs1=nothing
    rs.close
    set rs=nothing
    closedb()
  else
    rs1.close
    set rs1=nothing
    closedb()
  end if
  showctail
end sub

sub ShowInputForm201(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 202 then
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh", conn, 1, 1
    %>
    <form method="POST" action="marea-tjbb.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh", conn, 1, 1
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from tjlb where tjlb_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-tjbb.asp?mode=203&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cpage3%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cpage3%>">����</a>]
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
              <% if mode = 202 then%>
                ������С�࣬Ȼ������OK��
              <%else%>
                ��༭С�࣬Ȼ������OK��
                <input name="odq" type="hidden" value="<%=request("odq")%>">
              <%end if%>
            <%end if%>
            <input name="page1" type="hidden" value="<%=cpage1%>">
            <input name="dqcode1" type="hidden" value="<%=request("dqcode1")%>">
            <input name="dqname1" type="hidden" value="<%=request("dqname1")%>">
            <input name="page2" type="hidden" value="<%=cpage2%>">
            <input name="dqcode2" type="hidden" value="<%=request("dqcode2")%>">
            <input name="dqname2" type="hidden" value="<%=request("dqname2")%>">
            <input name="gs_cc2" type="hidden" value="<%=request("gs_cc2")%>">
            <input name="gs2" type="hidden" value="<%=request("gs2")%>">
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>С�����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=dq size=15 maxlength=<%=tjlb_len2-tjlb_len1%> class="smallInput" value='<%=request("dq")%>'>(ǰ<%=tjlb_len2%>λΪ<%=left(request("dqcode2"),tjlb_len2)%>,�����<%=tjlb_len3%>λ)
            <%else%>
              <input name=dq size=15 maxlength=<%=tjlb_len2-tjlb_len1%> class="smallInput" value='<%=trim(mid(rs("tjlb_code"),tjlb_len2+1,tjlb_len3-tjlb_len2))%>'>(ǰ<%=tjlb_len2%>λΪ<%=left(request("dqcode2"),tjlb_len2)%>,�����<%=tjlb_len3%>λ)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>С������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("tjlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�Ƿ���ʾ&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <input type="checkbox" name="sfxsxj" value='yes'>
            <%else%>
              <%if rs("sfxsxj")="Y" then%>
                <input type="checkbox" name="sfxsxj" value='yes' checked>
              <%else%>
                <input type="checkbox" name="sfxsxj" value='yes'>
              <%end if%>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��Ŀ����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=202 then%>
              <select name="FRMbz" OnClick="javascript:if(document.input1.FRMbz.value*1==0){test1.style.display=''}">
                <option value="0">��ͨ</option>
              </select>
            <%else%>
              <select name="FRMbz" OnClick="javascript:if(document.input1.FRMbz.value*1==0){test1.style.display=''}">
                <option value="0"<%if rs("bz")="-" then%>selected<%end if%>>��ͨ</option>
              </select>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��ʽ&nbsp;</td>
          <td align=left colspan=2>
             <%if mode=202 then%>
               <DIV ID="test1" Style="position:relative; display:'';"> 
               ������������:
               <br>���
               <select name="ajlb_code0" style="HEIGHT:17px;WIDTH:119px">
                 <option value="0"<%if gs_cc="0" then %> selected <% end if %>></option>
                 <%if request("gs_cc2")<1 then%>
                   <option value="1"<%if gs_cc="1" then %> selected <% end if %>>����</option>
                 <%end if%>
                 <%if request("gs_cc2")<2 then%>
                   <option value="2"<%if gs_cc="2" then %> selected <% end if %>>����</option>
                 <%end if%>
                 <%if request("gs_cc2")<3 then%>
                   <option value="3"<%if gs_cc="3" then %> selected <% end if %>>С��</option>
                 <%end if%>
                 <%if request("gs_cc2")<4 then%>
                   <option value="4"<%if gs_cc="4" then %> selected <% end if %>>�������</option>
                 <%end if%>
               </select>
               <br>����
               <%if request("gs_cc2")>=1 then%>
                 <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" disabled>
               <%else%>
                 <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(202,1)">
               <%end if%>
               <%while not rs1.EOF 
                 if trim(ajlb_code)="" then ajlb_code=trim(rs1("ajlb_code"))%>
                 <option value="<%=trim(rs1("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len1)=left(rs1("ajlb_code"),ajlb_len1) then %> selected <% end if %>><%=trim(rs1("ajlb_name"))%></option>
                 <%rs1.MoveNext 
               WEND%>
               </select>
               <br>����
               <%if request("gs_cc2")>=2 then%>
                 <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" disabled>
               <%else%>
                 <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(202,2)">
               <%end if%>
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(ajlb_code,ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF 
                 if mid(ajlb_code,ajlb_len1+1,ajlb_len2-ajlb_len1)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len2)=left(rsMX("ajlb_code"),ajlb_len2) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>С��
               <%if request("gs_cc2")>=3 then%>
                 <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" disabled>
               <%else%>
                 <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(202,3)">
               <%end if%>
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(ajlb_code,ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF 
                 if mid(ajlb_code,ajlb_len2+1,ajlb_len3-ajlb_len2)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len3)=left(rsMX("ajlb_code"),ajlb_len3) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>�������
               <select name="ajlb_code4" style="HEIGHT:17px;WIDTH:119px" >
               <option value="" <%if fxlb_code="" then %> selected <% end if %>></option>
               <%rsMX.open "select * from fxlb where left(fxlb_code," & ajlb_len3 & ")='" & left(ajlb_code,ajlb_len3) &"' order by fxlb_sxh", conn, 1, 1
               while not rsMX.EOF%>
                 <option value="<%=trim(rsMX("fxlb_code"))%>"<%if left(fxlb_code,fxlb_len1)=left(rsMX("fxlb_code"),fxlb_len1) then %> selected <% end if %>><%=trim(rsMX("fxlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               </DIV>
             <%else%>
               <%if rs("bz")="-" or rs("bz")="QNTQ" then%>
                 <DIV ID="test1" Style="position:relative; display:'';"> 
               <%else%>
                 <DIV ID="test1" Style="position:relative; display:'none';"> 
               <%end if%>
               ������������
               <br>���
               <select name="ajlb_code0" style="HEIGHT:17px;WIDTH:119px">
                 <option value="0"<%if gs_cc="0" then %> selected <% end if %>></option>
                 <%if request("gs_cc2")<1 then%>
                   <option value="1"<%if gs_cc="1" then %> selected <% end if %>>����</option>
                 <%end if%>
                 <%if request("gs_cc2")<2 then%>
                   <option value="2"<%if gs_cc="2" then %> selected <% end if %>>����</option>
                 <%end if%>
                 <%if request("gs_cc2")<3 then%>
                   <option value="3"<%if gs_cc="3" then %> selected <% end if %>>С��</option>
                 <%end if%>
                 <%if request("gs_cc2")<4 then%>
                   <option value="4"<%if gs_cc="4" then %> selected <% end if %>>�������</option>
                 <%end if%>
               </select>
               <br>����
               <%if request("gs_cc2")>=1 then%>
                 <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" disabled>
               <%else%>
                 <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(203,1)">
               <%end if%>
               <%while not rs1.EOF
                 if trim(ajlb_code)="" then ajlb_code=trim(rs1("ajlb_code"))%>
                 <option value="<%=trim(rs1("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len1)=left(rs1("ajlb_code"),ajlb_len1) then %> selected <% end if %>><%=trim(rs1("ajlb_name"))%></option>
                 <%rs1.MoveNext 
               WEND%>
               </select>
               <br>����
               <%if request("gs_cc2")>=2 then%>
                 <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" disabled>
               <%else%>
                 <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(203,2)">
               <%end if%>
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(ajlb_code,ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF
                 if mid(ajlb_code,ajlb_len1+1,ajlb_len2-ajlb_len1)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len2)=left(rsMX("ajlb_code"),ajlb_len2) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>С��
               <%if request("gs_cc2")>=3 then%>
                 <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" disabled>
               <%else%>
                 <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(203,3)">
               <%end if%>
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(ajlb_code,ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF
                 if mid(ajlb_code,ajlb_len2+1,ajlb_len3-ajlb_len2)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len3)=left(rsMX("ajlb_code"),ajlb_len3) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>�������
               <select name="ajlb_code4" style="HEIGHT:17px;WIDTH:119px" >
               <option value="" <%if ajlb_code="" then %> selected <% end if %>></option>
               <%rsMX.open "select * from fxlb where left(fxlb_code," & ajlb_len3 & ")='" & left(ajlb_code,ajlb_len3) &"' order by fxlb_sxh", conn, 1, 1
               while not rsMX.EOF%>
                 <option value="<%=trim(rsMX("fxlb_code"))%>"<%if left(ajlb_code,fxlb_len1)=left(rsMX("fxlb_code"),fxlb_len1) then %> selected <% end if %>><%=trim(rsMX("fxlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               </DIV>
             <%end if%>
           </td>
        </tr>
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
  if mode = 103 then
    rs1.close
    set rs1=nothing
    rs.close
    set rs=nothing
    closedb()
  else
    rs1.close
    set rs1=nothing
    closedb()
  end if
  showctail
end sub

sub ShowInputForm301(mode,errmsg)
  'on error resume next
  showchead()

  if mode = 302 then
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh", conn, 1, 1
    %>
    <form method="POST" action="marea-tjbb.asp?mode=302&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&odq=<%=request("odq")%>" name="input1">
  <%else
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    rs1.open "select * from ajlb where right(ajlb_code,"& (ajlb_len0-ajlb_len1) & ")='" & ajlb_str1 &"' order by ajlb_sxh", conn, 1, 1
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from tjlb where tjlb_code='" + request("odq") + "'", conn, 1, 1
    %>
    <form method="POST" action="marea-tjbb.asp?mode=303&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page4=<%=cpage4%>&odq=<%=request("odq")%>" name="input1">
  <%end if%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="marea-tjbb.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page4=<%=cpage4%>">����</a>]
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
              <% if mode = 302 then%>
                ������С1�࣬Ȼ������OK��
              <%else%>
                ��༭С1�࣬Ȼ������OK��
                <input name="odq" type="hidden" value="<%=request("odq")%>">
              <%end if%>
            <%end if%>
            <input name="page1" type="hidden" value="<%=cpage1%>">
            <input name="dqcode1" type="hidden" value="<%=request("dqcode1")%>">
            <input name="dqname1" type="hidden" value="<%=request("dqname1")%>">
            <input name="page2" type="hidden" value="<%=cpage2%>">
            <input name="dqcode2" type="hidden" value="<%=request("dqcode2")%>">
            <input name="dqname2" type="hidden" value="<%=request("dqname2")%>">
            <input name="page3" type="hidden" value="<%=cpage3%>">
            <input name="dqcode3" type="hidden" value="<%=request("dqcode3")%>">
            <input name="dqname3" type="hidden" value="<%=request("dqname3")%>">
            <input name="gs_cc2" type="hidden" value="<%=request("gs_cc2")%>">
            <input name="gs2" type="hidden" value="<%=request("gs2")%>">
            <input name="gs_cc3" type="hidden" value="<%=request("gs_cc3")%>">
            <input name="gs3" type="hidden" value="<%=request("gs3")%>">
          </td>
        </tr>
        <tr><td colspan="3"><hr noshade size="1" width="100%"></td></tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>С1�����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=302 then%>
              <input name=dq size=15 maxlength=<%=tjlb_len2-tjlb_len1%> class="smallInput" value='<%=request("dq")%>'>(ǰ<%=tjlb_len3%>λΪ<%=left(request("dqcode3"),tjlb_len3)%>,�����<%=tjlb_len4-tjlb_len3%>λ)
            <%else%>
              <input name=dq size=15 maxlength=<%=tjlb_len2-tjlb_len1%> class="smallInput" value='<%=trim(mid(rs("tjlb_code"),tjlb_len3+1,tjlb_len4-tjlb_len3))%>'>(ǰ<%=tjlb_len3%>λΪ<%=left(request("dqcode3"),tjlb_len3)%>,�����<%=tjlb_len4-tjlb_len3%>λ)
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>С1������&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=302 then%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=request("dq0")%>'>
            <%else%>
              <input name=dq0 size=25 maxlength=30 class="smallInput" value='<%=trim(rs("tjlb_name"))%>'>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>�Ƿ���ʾ&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=302 then%>
              <input type="checkbox" name="sfxsxj" value='yes'>
            <%else%>
              <%if rs("sfxsxj")="Y" then%>
                <input type="checkbox" name="sfxsxj" value='yes' checked>
              <%else%>
                <input type="checkbox" name="sfxsxj" value='yes'>
              <%end if%>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��Ŀ����&nbsp;</td>
          <td align=left colspan=2>
            <%if mode=302 then%>
              <select name="FRMbz" OnClick="javascript:if(document.input1.FRMbz.value*1==0){test1.style.display=''}">
                <option value="0">��ͨ</option>
              </select>
            <%else%>
              <select name="FRMbz" OnClick="javascript:if(document.input1.FRMbz.value*1==0){test1.style.display=''}">
                <option value="0"<%if rs("bz")="-" then%>selected<%end if%>>��ͨ</option>
              </select>
            <%end if%>
          </td>
        </tr>
        <tr>
          <td bgcolor="#eeeeee" align=right nowrap width=20%>��ʽ&nbsp;</td>
          <td align=left colspan=2>
             <%if mode=302 then%>
               <DIV ID="test1" Style="position:relative; display:'';"> 
               ������������:
               <br>���
               <select name="ajlb_code0" style="HEIGHT:17px;WIDTH:119px">
                 <%if request("gs_cc3")=0 then%>
                   <%if request("gs_cc2")<2 then%>
                     <option value="2"<%if gs_cc="2" then %> selected <% end if %>>����</option>
                   <%end if%>
                   <%if request("gs_cc2")<3 then%>
                     <option value="3"<%if gs_cc="3" then %> selected <% end if %>>С��</option>
                   <%end if%>
                   <%if request("gs_cc2")<4 then%>
                     <option value="4"<%if gs_cc="4" then %> selected <% end if %>>�������</option>
                   <%end if%>
                 <%else%>
                   <%if request("gs_cc3")<2 then%>
                     <option value="2"<%if gs_cc="2" then %> selected <% end if %>>����</option>
                   <%end if%>
                   <%if request("gs_cc3")<3 then%>
                     <option value="3"<%if gs_cc="3" then %> selected <% end if %>>С��</option>
                   <%end if%>
                   <%if request("gs_cc3")<4 then%>
                     <option value="4"<%if gs_cc="4" then %> selected <% end if %>>�������</option>
                   <%end if%>
                 <%end if%>
               </select>
               <br>����
               <%if request("gs_cc3")=0 then%>
                 <%if request("gs_cc2")>=1 then%>
                   <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,1)">
                 <%end if%>
               <%else%>
                 <%if request("gs_cc3")>=1 then%>
                   <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,1)">
                 <%end if%>
               <%end if%>
               <%while not rs1.EOF 
                 if trim(ajlb_code)="" then ajlb_code=trim(rs1("ajlb_code"))%>
                 <option value="<%=trim(rs1("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len1)=left(rs1("ajlb_code"),ajlb_len1) then %> selected <% end if %>><%=trim(rs1("ajlb_name"))%></option>
                 <%rs1.MoveNext 
               WEND%>
               </select>
               <br>����
               <%if request("gs_cc3")=0 then%>
                 <%if request("gs_cc2")>=2 then%>
                   <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,2)">
                 <%end if%>
               <%else%>
                 <%if request("gs_cc3")>=2 then%>
                   <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,2)">
                 <%end if%>
               <%end if%>
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(ajlb_code,ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF 
                 if mid(ajlb_code,ajlb_len1+1,ajlb_len2-ajlb_len1)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len2)=left(rsMX("ajlb_code"),ajlb_len2) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>С��
               <%if request("gs_cc3")=0 then%>
                 <%if request("gs_cc2")>=3 then%>
                   <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,3)">
                 <%end if%>
               <%else%>
                 <%if request("gs_cc3")>=3 then%>
                   <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,3)">
                 <%end if%>
               <%end if%>
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(ajlb_code,ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF 
                 if mid(ajlb_code,ajlb_len2+1,ajlb_len3-ajlb_len2)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len3)=left(rsMX("ajlb_code"),ajlb_len3) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>�������
               <select name="ajlb_code4" style="HEIGHT:17px;WIDTH:119px" >
               <option value="" <%if fxlb_code="" then %> selected <% end if %>></option>
               <%rsMX.open "select * from fxlb where left(fxlb_code," & ajlb_len3 & ")='" & left(ajlb_code,ajlb_len3) &"' order by fxlb_sxh", conn, 1, 1
               while not rsMX.EOF%>
                 <option value="<%=trim(rsMX("fxlb_code"))%>"<%if left(fxlb_code,fxlb_len1)=left(rsMX("fxlb_code"),fxlb_len1) then %> selected <% end if %>><%=trim(rsMX("fxlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               </DIV>
             <%else%>
               <%if rs("bz")="-" or rs("bz")="QNTQ" then%>
                 <DIV ID="test1" Style="position:relative; display:'';"> 
               <%else%>
                 <DIV ID="test1" Style="position:relative; display:'none';"> 
               <%end if%>
               ������������
               <br>���
               <select name="ajlb_code0" style="HEIGHT:17px;WIDTH:119px">
                 <%if request("gs_cc3")=0 then%>
                   <%if request("gs_cc2")<2 then%>
                     <option value="2"<%if gs_cc="2" then %> selected <% end if %>>����</option>
                   <%end if%>
                   <%if request("gs_cc2")<3 then%>
                     <option value="3"<%if gs_cc="3" then %> selected <% end if %>>С��</option>
                   <%end if%>
                   <%if request("gs_cc2")<4 then%>
                     <option value="4"<%if gs_cc="4" then %> selected <% end if %>>�������</option>
                   <%end if%>
                 <%else%>
                   <%if request("gs_cc3")<2 then%>
                     <option value="2"<%if gs_cc="2" then %> selected <% end if %>>����</option>
                   <%end if%>
                   <%if request("gs_cc3")<3 then%>
                     <option value="3"<%if gs_cc="3" then %> selected <% end if %>>С��</option>
                   <%end if%>
                   <%if request("gs_cc3")<4 then%>
                     <option value="4"<%if gs_cc="4" then %> selected <% end if %>>�������</option>
                   <%end if%>
                 <%end if%>
               </select>
               <br>����
               <%if request("gs_cc3")=0 then%>
                 <%if request("gs_cc2")>=1 then%>
                   <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,1)">
                 <%end if%>
               <%else%>
                 <%if request("gs_cc3")>=1 then%>
                   <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code1" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,1)">
                 <%end if%>
               <%end if%>
               <%while not rs1.EOF
                 if trim(ajlb_code)="" then ajlb_code=trim(rs1("ajlb_code"))%>
                 <option value="<%=trim(rs1("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len1)=left(rs1("ajlb_code"),ajlb_len1) then %> selected <% end if %>><%=trim(rs1("ajlb_name"))%></option>
                 <%rs1.MoveNext 
               WEND%>
               </select>
               <br>����
               <%if request("gs_cc3")=0 then%>
                 <%if request("gs_cc2")>=2 then%>
                   <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,2)">
                 <%end if%>
               <%else%>
                 <%if request("gs_cc3")>=1 then%>
                   <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code2" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,2)">
                 <%end if%>
               <%end if%>
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len1 & ")='" & left(ajlb_code,ajlb_len1) &"' and right(ajlb_code,"& (ajlb_len0-ajlb_len2) & ")='" & ajlb_str2 &"' and mid(ajlb_code,"& (ajlb_len1+1) & "," & (ajlb_len2-ajlb_len1) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF
                 if mid(ajlb_code,ajlb_len1+1,ajlb_len2-ajlb_len1)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len2)=left(rsMX("ajlb_code"),ajlb_len2) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>С��
               <%if request("gs_cc3")=0 then%>
                 <%if request("gs_cc2")>=3 then%>
                   <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,3)">
                 <%end if%>
               <%else%>
                 <%if request("gs_cc3")>=3 then%>
                   <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" disabled>
                 <%else%>
                   <select name="ajlb_code3" style="HEIGHT:17px;WIDTH:119px" onchange="Getseconditem(302,3)">
                 <%end if%>
               <%end if%>
               <%rsMX.open "select * from ajlb where left(ajlb_code," & ajlb_len2 & ")='" & left(ajlb_code,ajlb_len2) &"' and mid(ajlb_code,"& (ajlb_len2+1) & "," & (ajlb_len3-ajlb_len2) & ")<>'00' order by ajlb_sxh", conn, 1, 1
               while not rsMX.EOF
                 if mid(ajlb_code,ajlb_len2+1,ajlb_len3-ajlb_len2)="00" then ajlb_code=trim(rsMX("ajlb_code"))%>
                 <option value="<%=trim(rsMX("ajlb_code"))%>"<%if left(ajlb_code,ajlb_len3)=left(rsMX("ajlb_code"),ajlb_len3) then %> selected <% end if %>><%=trim(rsMX("ajlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               <br>�������
               <select name="ajlb_code4" style="HEIGHT:17px;WIDTH:119px" >
               <option value="" <%if ajlb_code="" then %> selected <% end if %>></option>
               <%rsMX.open "select * from fxlb where left(fxlb_code," & ajlb_len3 & ")='" & left(ajlb_code,ajlb_len3) &"' order by fxlb_sxh", conn, 1, 1
               while not rsMX.EOF%>
                 <option value="<%=trim(rsMX("fxlb_code"))%>"<%if left(ajlb_code,fxlb_len1)=left(rsMX("fxlb_code"),fxlb_len1) then %> selected <% end if %>><%=trim(rsMX("fxlb_name"))%></option>
                 <%rsMX.MoveNext 
               WEND
               rsMX.close%>
               </select>
               </DIV>
             <%end if%>
           </td>
        </tr>
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
  if mode = 103 then
    rs1.close
    set rs1=nothing
    rs.close
    set rs=nothing
    closedb()
  else
    rs1.close
    set rs1=nothing
    closedb()
  end if
  showctail
end sub

if mode=1 then
  '������ʾ 
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  'response.write "select * from tjlb where right(tjlb_code,"& (tjlb_len0-tjlb_len1) & ")='" & tjlb_str1 &"' order by tjlb_sxh"
  rs.open "select * from tjlb where right(tjlb_code,"& (tjlb_len0-tjlb_len1) & ")='" & tjlb_str1 &"' order by tjlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage1>rs.pagecount then cpage1=1
    rs.AbsolutePage = cpage1
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">��<%=cstr(cpage1)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>��һ����Ŀ</td>
          <td align="right">
          [<a href="marea-tjbb.asp?mode=2">���</a>]
          <%if cpage1 <> 1 then%>
            [<a href="marea-tjbb.asp?mode=1&page1=<%=cstr(cpage1-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage1 <> rs.PageCount then%>
            [<a href="marea-tjbb.asp?mode=1&page1=<%=cstr(cpage1+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage1 then%>
                <option selected value="marea-tjbb.asp?mode=1&page1=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-tjbb.asp?mode=1&page1=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
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
              <td width=40% align=center>��������</td>
              <td width=50% align=center>����</td>
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
                <td align=center><%=trim(rs("tjlb_code"))%></td>
                <td align=center><%=trim(rs("tjlb_name"))%><a href="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=trim(rs("tjlb_code"))%>&dqname1=<%=trim(rs("tjlb_name"))%>">��<font color="#FF0000">����</font>��</a></td>
                <td align=center>
                  <a href="marea-tjbb.asp?mode=3&page1=<%=cpage1%>&odq=<%=trim(rs("tjlb_code"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-tjbb.asp?mode=4&page1=<%=cpage1%>&dq=<%=trim(rs("tjlb_code"))%>&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("tjlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-tjbb.asp?mode=8&page1=<%=cpage1%>&dq=<%=trim(rs("tjlb_code"))%>&sort=up&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("tjlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-tjbb.asp?mode=8&page1=<%=cpage1%>&dq=<%=trim(rs("tjlb_code"))%>&sort=down&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
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
          [<a href="marea-tjbb.asp?mode=2">���</a>]
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
elseif mode=101 then
  '������ʾ
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  'Response.Write("select * from tjlb where left(tjlb_code," & tjlb_len1 & ")='" & left(request("dqcode1"),tjlb_len1) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len2) & ")='" & tjlb_str2 &"' and mid(tjlb_code,"& (tjlb_len1+1) & "," & (tjlb_len2-tjlb_len1) & ")<>'00' order by tjlb_sxh")
  rs.open "select * from tjlb where left(tjlb_code," & tjlb_len1 & ")='" & left(request("dqcode1"),tjlb_len1) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len2) & ")='" & tjlb_str2 &"' and mid(tjlb_code,"& (tjlb_len1+1) & "," & (tjlb_len2-tjlb_len1) & ")<>'00' order by tjlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage2>rs.pagecount then cpage2=1
    rs.AbsolutePage = cpage2
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">��<%=cstr(cpage2)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>��������Ŀ</td>
          <td align="right">
          [<a href="marea-tjbb.asp?mode=1&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">�����б�</a>]&nbsp;
          [<a href="marea-tjbb.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">���</a>]
          <%if cpage2 <> 1 then%>
            [<a href="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(cpage2-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage2 <> rs.PageCount then%>
            [<a href="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(cpage2+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2" onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage2 then%>
                <option selected value="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
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
              <td width=40% align=center>��������</td>
              <td width=50% align=center>����</td>
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
                <td align=center><%=trim(rs("tjlb_code"))%></td>
                <td align=center><%=trim(rs("tjlb_name"))%><%if rs("gs_cc")<4 then%><a href="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=trim(request("dqcode1"))%>&dqname1=<%=trim(request("dqname1"))%>&page2=<%=cpage2%>&dqcode2=<%=trim(rs("tjlb_code"))%>&dqname2=<%=trim(rs("tjlb_name"))%>&gs_cc2=<%=trim(rs("gs_cc"))%>&gs2=<%=rs("gs")%>">��<font color="#FF0000">С��</font>��</a><%end if%></td>
                <td align=center>
                  <a href="marea-tjbb.asp?mode=103&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&odq=<%=trim(rs("tjlb_code"))%>&ajlb_code=<%=rs("gs")%>&gs_cc=<%=rs("gs_cc")%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-tjbb.asp?mode=104&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("tjlb_code"))%>&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("tjlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-tjbb.asp?mode=108&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("tjlb_code"))%>&sort=up&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("tjlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-tjbb.asp?mode=108&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dq=<%=trim(rs("tjlb_code"))%>&sort=down&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
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
          [<a href="marea-tjbb.asp?mode=1&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">�����б�</a>]&nbsp;
          [<a href="marea-tjbb.asp?mode=102&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">���</a>]
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

elseif mode=201 then
  'С����ʾ
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  'rs.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(request("dqcode2"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh", conn, 1, 1
  rs.open "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(request("dqcode2"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh", conn, 1, 1
  'response.write "select * from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(request("dqcode2"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh"
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage3>rs.pagecount then cpage3=1
    rs.AbsolutePage = cpage3
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">��<%=cstr(cpage2)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>��С����Ŀ</td>
          <td align="right">
          [<a href="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>">�����б�</a>]&nbsp;
          [<a href="marea-tjbb.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&ajlb_code=<%=request("gs2")%>&gs_cc=<%=(request("gs_cc2")+1)%>">���</a>]
          <%if cpage3 <> 1 then%>
            [<a href="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cstr(cpage3-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage3 <> rs.PageCount then%>
            [<a href="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cstr(cpage3+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage3 then%>
                <option selected value="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>С�����</td>
              <td width=40% align=center>С������</td>
              <td width=50% align=center>����</td>
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
                <td align=center><%=trim(rs("tjlb_code"))%></td>
                <td align=center><%=trim(rs("tjlb_name"))%><%if rs("gs_cc")<4 then%><a href="marea-tjbb.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=trim(request("dqcode1"))%>&dqname1=<%=trim(request("dqname1"))%>&page2=<%=cpage2%>&dqcode2=<%=trim(request("dqcode2"))%>&dqname2=<%=trim(request("dqname2"))%>&page3=<%=cpage3%>&dqcode3=<%=trim(rs("tjlb_code"))%>&dqname3=<%=trim(rs("tjlb_name"))%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=trim(rs("gs_cc"))%>&gs3=<%=rs("gs")%>">��<font color="#FF0000">С1��</font>��</a><%end if%></td>
                <td align=center>
                  <%if request("gs_cc2")=1 and left(request("gs2"),ajlb_len1)<>left(rs("gs"),ajlb_len1) then
                    str="&ajlb_code=" &request("gs2")
                  elseif request("gs_cc2")=2 and left(request("gs2"),ajlb_len2)<>left(rs("gs"),ajlb_len2) then
                    str="&ajlb_code=" &request("gs2")
                  elseif request("gs_cc2")=3 and left(request("gs2"),ajlb_len3)<>left(rs("gs"),ajlb_len3) then
                    str="&ajlb_code=" &request("gs2")
                  else
                    str="&ajlb_code=" &rs("gs")
                  end if
                  if rs("gs_cc")=0 then
                    str=str & "&gs_cc=" &rs("gs_cc")
                  elseif request("gs_cc2")>rs("gs_cc") then
                    str=str & "&gs_cc=" &(request("gs_cc2")+1)
                  else
                    str=str & "&gs_cc=" &rs("gs_cc")
                  end if
                  %>
                  <a href="marea-tjbb.asp?mode=203&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cpage3%>&odq=<%=trim(rs("tjlb_code")) & str%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-tjbb.asp?mode=204&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("tjlb_code"))%>&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("tjlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-tjbb.asp?mode=208&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("tjlb_code"))%>&sort=up&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("tjlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-tjbb.asp?mode=208&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cpage3%>&dq=<%=trim(rs("tjlb_code"))%>&sort=down&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
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
          [<a href="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>">�����б�</a>]&nbsp;
          [<a href="marea-tjbb.asp?mode=202&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&ajlb_code=<%=request("gs2")%>&gs_cc=<%=(request("gs_cc2")+1)%>">���</a>]
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

elseif mode=301 then
  'С1����ʾ
  showchead()
  Response.Write "<br>"
  opendb()

  set rs=server.createobject("adodb.recordset")
  rs.open "select * from tjlb where left(tjlb_code," & tjlb_len3 & ")='" & left(request("dqcode3"),tjlb_len3) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len4) & ")='" & tjlb_str4 &"' and mid(tjlb_code,"& (tjlb_len3+1) & "," & (tjlb_len4-tjlb_len3) & ")<>'00' order by tjlb_sxh", conn, 1, 1
  if rs.recordcount <> 0 then
    rs.movefirst
    rs.CacheSize = 5
    rs.PageSize = 10
    if cpage3>rs.pagecount then cpage3=1
    rs.AbsolutePage = cpage3
    %>
      <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td valign="bottom">��<%=cstr(cpage2)%>ҳ/��<%=cstr(rs.PageCount)%>ҳ����<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>��С1����Ŀ</td>
          <td align="right">
          [<a href="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>">С���б�</a>]&nbsp;
          [<a href="marea-tjbb.asp?mode=302&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&ajlb_code=<%=request("gs2")%>&gs_cc=<%=(request("gs_cc3")+1)%>">���</a>]
          <%if cpage3 <> 1 then%>
            [<a href="marea-tjbb.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page3=<%=cstr(cpage3-1)%>">��һҳ</a>]
          <%end if%>
          <%if cpage3 <> rs.PageCount then%>
            [<a href="marea-tjbb.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page3=<%=cstr(cpage3+1)%>">��һҳ</a>]
          <%end if%>
          <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage3 then%>
                <option selected value="marea-tjbb.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page3=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%else%>
                <option value="marea-tjbb.asp?mode=301&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page3=<%=cstr(i)%>">���� <%=cstr(i)%> ҳ</option>
              <%end if%>
             <%next%>
          </select>
          <%end if%>
          </td>
        </tr>
        <tr><td colspan="2">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
              <td width=10% align=center>С1�����</td>
              <td width=40% align=center>С1������</td>
              <td width=50% align=center>����</td>
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
                <td align=center><%=trim(rs("tjlb_code"))%></td>
                <td align=center><%=trim(rs("tjlb_name"))%></td>
                <td align=center>
                  <%if request("gs_cc3")=1 and left(request("gs3"),ajlb_len1)<>left(rs("gs"),ajlb_len1) then
                    str="&ajlb_code=" &request("gs2")
                  elseif request("gs_cc3")=2 and left(request("gs3"),ajlb_len2)<>left(rs("gs"),ajlb_len2) then
                    str="&ajlb_code=" &request("gs2")
                  elseif request("gs_cc3")=3 and left(request("gs3"),ajlb_len3)<>left(rs("gs"),ajlb_len3) then
                    str="&ajlb_code=" &request("gs2")
                  else
                    str="&ajlb_code=" &rs("gs")
                  end if
                  if request("gs_cc3")>rs("gs_cc") then
                    str=str & "&gs_cc=" &(request("gs_cc3")+1)
                  else
                    str=str & "&gs_cc=" &rs("gs_cc")
                  end if
                  %>
                  <a href="marea-tjbb.asp?mode=303&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page4=<%=cpage4%>&odq=<%=trim(rs("tjlb_code")) & str%>"><img src="./images/edit.gif" border=0></a>
                  <a href="marea-tjbb.asp?mode=304&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page4=<%=cpage4%>&dq=<%=trim(rs("tjlb_code"))%>&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/del.gif" border=0></a>
                  <%if rs("tjlb_sxh")=1 then%>
                    <img src="./images/up.gif" border=0>
                  <%else%>
                    <a href="marea-tjbb.asp?mode=308&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page4=<%=cpage4%>&dq=<%=trim(rs("tjlb_code"))%>&sort=up&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/up.gif" border=0></a>
                  <%end if%>
                  <%if rs("tjlb_sxh")=rs.RecordCount then%>
                    <img src="./images/down.gif" border=0>
                  <%else%>
                    <a href="marea-tjbb.asp?mode=308&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&page4=<%=cpage4%>&dq=<%=trim(rs("tjlb_code"))%>&sort=down&dwxh=<%=trim(rs("tjlb_sxh"))%>"><img src="./images/down.gif" border=0></a>
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
          [<a href="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>">С���б�</a>]&nbsp;
          [<a href="marea-tjbb.asp?mode=302&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&page3=<%=cpage3%>&dqcode3=<%=request("dqcode3")%>&dqname3=<%=request("dqname3")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&gs_cc3=<%=request("gs_cc3")%>&gs3=<%=request("gs3")%>&ajlb_code=<%=request("gs2")%>&gs_cc=<%=(request("gs_cc2")+1)%>">���</a>]
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

elseif mode=2 or mode=3 then
  '������Ӽ��޸�
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to tjlb_len1-1
      dq="0"+cstr(dq)
    next 
    for i=len(dq) to tjlb_len0-1
      dq=cstr(dq)+"0"
    next 
    dq0 = trim(request("dq0"))
    if request.form("sfxsxj")="" then'�Ƿ���ʾ
      sfxsxj="N"
    else
      sfxsxj="Y"
    end if 
    'response.write dq
    if mode=2 then
      if dq = "" then
        ErrMsg="������������"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select tjlb_name from tjlb where tjlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "��������ظ�"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select * from tjlb where right(tjlb_code,"& (tjlb_len0-tjlb_len1) & ")='" & tjlb_str1 &"' order by tjlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="�������������"
        foundError=True
      end if
    else
      '���Ĺ����û����Ƿ����
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select tjlb_name from tjlb where tjlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "��������ظ�"
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
        '�����
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "tjlb", conn, 1, 3
        rs.addnew
        rs("tjlb_code")=dq
        rs("tjlb_name")=dq0
        rs("tjlb_sxh")=dwxh
        rs("sfxsxj")=sfxsxj
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-tjbb.asp?mode=1"
      else
        opendb()
        conn.Execute "update tjlb set tjlb_code='"+dq+"',tjlb_name='"+dq0+"',sfxsxj='"+sfxsxj+"' where tjlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-tjbb.asp?mode=1&page1=" & cpage1
      end if
    end if
  else
      ShowInputForm1 mode,""
  end if

elseif mode=102 or mode=103 then
  '������Ӽ��޸�
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to tjlb_len2-tjlb_len1-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode1"),tjlb_len1)+ dq
    for i=len(dq) to tjlb_len0-1
      dq=cstr(dq)+"0"
    next
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    gs_cc=request("ajlb_code0")
    if request.form("sfxsxj")="" then'�Ƿ���ʾ
      sfxsxj="N"
    else
      sfxsxj="Y"
    end if 
    if request("FRMbz")=0 then'����
      dqbz="-"
      if gs_cc=1 then
        dqgs=request("ajlb_code1")
      elseif gs_cc=2 then
        dqgs=request("ajlb_code2")
      elseif gs_cc=3 then
        dqgs=request("ajlb_code3")
      elseif gs_cc=4 then
        dqgs=request("ajlb_code4")
      end if
    elseif request("FRMbz")=1 then'����
      dqbz="QT"
      dqgs=""
    end if
    if mode=102 then
      if dq = "" then
        ErrMsg="�������������"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select tjlb_name from tjlb where tjlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "��������ظ�"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select tjlb_name from tjlb where left(tjlb_code," & tjlb_len1 & ")='" & left(request("dqcode1"),tjlb_len1) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len2) & ")='" & tjlb_str2 &"' and mid(tjlb_code,"& (tjlb_len1+1) & "," & (tjlb_len2-tjlb_len1) & ")<>'00' order by tjlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="��������������"
        foundError=True
      end if
    else
      '���Ĺ����û����Ƿ����
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select tjlb_name from tjlb where tjlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "��������ظ�"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm101 mode,errmsg
    else
      if mode = 102 then
        '�����
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "tjlb", conn, 1, 3
        rs.addnew
        rs("tjlb_code")=dq
        rs("tjlb_name")=dq0
        rs("tjlb_sxh")=dwxh
        rs("bz")=dqbz
        rs("gs")=dqgs
        rs("gs_cc")=gs_cc
        rs("sfxsxj")=sfxsxj
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-tjbb.asp?mode=101&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")
      else
        opendb()
        conn.Execute "update tjlb set tjlb_code='"+dq+"',tjlb_name='"+dq0+"',bz='"+dqbz+"',gs='"+dqgs+"',gs_cc='"+gs_cc+"',sfxsxj='"+sfxsxj+"' where tjlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-tjbb.asp?mode=101&page1="& cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2
      end if
    end if
  else
      ShowInputForm101 mode,""
  end if

elseif mode=202 or mode=203 then
  'С����Ӽ��޸�
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to tjlb_len3-tjlb_len2-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode2"),tjlb_len2)+ dq
    for i=len(dq) to tjlb_len0-1
      dq=cstr(dq)+"0"
    next
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    if request.form("sfxsxj")="" then'�Ƿ���ʾ
      sfxsxj="N"
    else
      sfxsxj="Y"
    end if    
    gs_cc=request("ajlb_code0")
    'response.write gs_cc
    if gs_cc=0 then
    else
      if gs_cc<request("gs_cc2") then gs_cc=request("gs_cc2")+1
    end if
    'response.write gs_cc
    if request("FRMbz")=0 then'����
      dqbz="-"
      if gs_cc=0 then
        dqgs=""
      elseif gs_cc=1 then
        dqgs=request("ajlb_code1")
      elseif gs_cc=2 then
        dqgs=request("ajlb_code2")
      elseif gs_cc=3 then
        dqgs=request("ajlb_code3")
      elseif gs_cc=4 then
        dqgs=request("ajlb_code4")
      end if
    end if
    'response.write dqgs
    if mode=202 then
      if dq = "" then
        ErrMsg="������С�����"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select tjlb_name from tjlb where tjlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "С������ظ�"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          'rs.open "select tjlb_name from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(request("dqcode2"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh", conn, 1, 1
          rs.open "select tjlb_name from tjlb where left(tjlb_code," & tjlb_len2 & ")='" & left(request("dqcode2"),tjlb_len2) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len3) & ")='" & tjlb_str3 &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' order by tjlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="������С������"
        foundError=True
      end if
    else
      '���Ĺ����û����Ƿ����
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select tjlb_name from tjlb where tjlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "С������ظ�"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm201 mode,errmsg
    else
      if mode = 202 then
        '�����
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "tjlb", conn, 1, 3
        rs.addnew
        rs("tjlb_code")=dq
        rs("tjlb_name")=dq0
        rs("tjlb_sxh")=dwxh
        rs("bz")=dqbz
        rs("gs")=dqgs
        rs("gs_cc")=gs_cc
        rs("sfxsxj")=sfxsxj
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-tjbb.asp?mode=201&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")& "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&gs_cc2=" & request("gs_cc2") & "&gs2=" & request("gs2")
      else
        opendb()
        conn.Execute "update tjlb set tjlb_code='"+dq+"',tjlb_name='"+dq0+"',bz='"+dqbz+"',gs='"+dqgs+"',gs_cc='"+gs_cc+"',sfxsxj='"+sfxsxj+"' where tjlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-tjbb.asp?mode=201&page1="& cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&gs_cc2=" & request("gs_cc2") & "&gs2=" & request("gs2") & "&page3=" & cpage3
      end if
    end if
  else
      ShowInputForm201 mode,""
  end if

elseif mode=302 or mode=303 then
  'С1����Ӽ��޸�
  if request("dq")<>"" and request("dq0")<>"" then
    FoundError=false
    ErrMsg=""
    dq =trim(request("dq"))
    for i=len(dq) to tjlb_len4-tjlb_len3-1
      dq="0"+cstr(dq)
    next
    dq =left(request("dqcode3"),tjlb_len3)+ dq
    for i=len(dq) to tjlb_len0-1
      dq=cstr(dq)+"0"
    next
    'response.write dq
    dq0 = trim(request("dq0"))
    dq1=trim(request("dq1"))
    if request.form("sfxsxj")="" then'�Ƿ���ʾ
      sfxsxj="N"
    else
      sfxsxj="Y"
    end if    
    gs_cc=request("ajlb_code0")
    'response.write gs_cc
    if gs_cc<request("gs_cc3") then gs_cc=request("gs_cc3")+1
    'response.write gs_cc
    if request("FRMbz")=0 then'����
      dqbz="-"
      if gs_cc=1 then
        dqgs=request("ajlb_code1")
      elseif gs_cc=2 then
        dqgs=request("ajlb_code2")
      elseif gs_cc=3 then
        dqgs=request("ajlb_code3")
      elseif gs_cc=4 then
        dqgs=request("ajlb_code4")
      end if
    end if
    'response.write dqgs
    if mode=302 then
      if dq = "" then
        ErrMsg="������С1�����"
        foundError=True
      else
        opendb()
        set rs=server.createobject("adodb.recordset")
        '�����Ƿ����ظ���ע��
        rs.open "select tjlb_name from tjlb where tjlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "С1������ظ�"
          FoundError = True
        end if
        rs.close
        if FoundError = false then
          rs.open "select tjlb_name from tjlb where left(tjlb_code," & tjlb_len3 & ")='" & left(request("dqcode3"),tjlb_len3) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len4) & ")='" & tjlb_str4 &"' and mid(tjlb_code,"& (tjlb_len3+1) & "," & (tjlb_len4-tjlb_len3) & ")<>'00' order by tjlb_sxh", conn, 1, 1
          dwxh=rs.RecordCount+1
          rs.close
        end if
        set rs=nothing
        closedb()
      end if
      if dq0 = "" then
        ErrMsg="������С1������"
        foundError=True
      end if
    else
      '���Ĺ����û����Ƿ����
      if request("odq")<>dq then
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "select tjlb_name from tjlb where tjlb_code='" + dq + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "С1������ظ�"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      end if
    end if
    if FoundError=true then
      ShowInputForm301 mode,errmsg
    else
      if mode = 302 then
        '�����
        opendb()
        set rs=server.createobject("adodb.recordset")
        rs.open "tjlb", conn, 1, 3
        rs.addnew
        rs("tjlb_code")=dq
        rs("tjlb_name")=dq0
        rs("tjlb_sxh")=dwxh
        rs("bz")=dqbz
        rs("gs")=dqgs
        rs("gs_cc")=gs_cc
        rs("sfxsxj")=sfxsxj
        rs.update
        rs.close
        set rs=nothing
        closedb()
        Response.Redirect "marea-tjbb.asp?mode=301&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")& "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2")& "&page3=" & cpage3& "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3") & "&gs_cc2=" & request("gs_cc2") & "&gs2=" & request("gs2")& "&gs_cc3=" & request("gs_cc3") & "&gs3=" & request("gs3")
      else
        opendb()
        conn.Execute "update tjlb set tjlb_code='"+dq+"',tjlb_name='"+dq0+"',bz='"+dqbz+"',gs='"+dqgs+"',gs_cc='"+gs_cc+"',sfxsxj='"+sfxsxj+"' where tjlb_code='"+request("odq")+"'"
        'update other table
        'conn.Execute "update bgk set dq='"+dq+"' where dq='"+request("odq")+"'"
        closedb()
        Response.Redirect "marea-tjbb.asp?mode=301&page1="& cpage1 & "&dqcode1="+request("dqcode1")  & "&dqname1=" & request("dqname1")& "&page2=" & cpage2& "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2")& "&page3=" & cpage3& "&dqcode3="+request("dqcode3") & "&dqname3=" & request("dqname3") & "&gs_cc2=" & request("gs_cc2") & "&gs2=" & request("gs2")& "&gs_cc3=" & request("gs_cc3") & "&gs3=" & request("gs3") & "&page4=" & cpage4
      end if
    end if
  else
      ShowInputForm301 mode,""
  end if

elseif mode=4 then
  '����ɾ��ȷ��
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      ���Ҫɾ�����ࡰ<%=request("dq")%>����
      <br><br>
      [<a href="marea-tjbb.asp?mode=7&page1=<%=cpage1%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-tjbb.asp?mode=1&page1=<%=cpage1%>">����</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=104 then
  '����ɾ��ȷ��
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      ���Ҫɾ�����ࡰ<%=request("dq")%>����
      <br><br>
      [<a href="marea-tjbb.asp?mode=107&page1=<%=cpage1%>&page2=<%=cpage2%>&dq=<%=request("dq")%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&dwxh=<%=request("dwxh")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-tjbb.asp?mode=101&page1=<%=cpage1%>&page2=<%=cpage2%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>">����</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=204 then
  'С��ɾ��ȷ��
  showchead()
  %>
  <br>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      ���Ҫɾ��С�ࡰ<%=request("dq")%>����
      <br><br>
      [<a href="marea-tjbb.asp?mode=207&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cpage3%>&dq=<%=request("dq")%>&dwxh=<%=request("dwxh")%>">�ǵ�</a>]
      &nbsp;&nbsp;&nbsp;[<a href="marea-tjbb.asp?mode=201&page1=<%=cpage1%>&dqcode1=<%=request("dqcode1")%>&dqname1=<%=request("dqname1")%>&page2=<%=cpage2%>&dqcode2=<%=request("dqcode2")%>&dqname2=<%=request("dqname2")%>&gs_cc2=<%=request("gs_cc2")%>&gs2=<%=request("gs2")%>&page3=<%=cpage3%>">����</a>]
      <br><br>
    </td></tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
  </table>
  <%
  showctail()

elseif mode=7 then
  '����delete
  opendb()
  conn.execute "delete from tjlb where tjlb_code like '" + left(request("dq"),tjlb_len1)+"%'"'���������༰�����������С��
  conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where right(tjlb_code,"& (tjlb_len0-tjlb_len1) & ")='" & tjlb_str1 &"' and tjlb_sxh>=" & request("dwxh")' �����˳�����ǰ��
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-tjbb.asp?mode=1&page1=" & cpage1)

elseif mode=107 then
  '����delete
  opendb()
  conn.execute "delete from tjlb where tjlb_code like'" + left(request("dq"),tjlb_len2)+"%'"'����������༰������С��
  conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where left(tjlb_code," & tjlb_len1 & ")='" & left(request("dqcode1"),tjlb_len1) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len2) & ")='" & tjlb_str2 &"' and mid(tjlb_code,"& (tjlb_len1+1) & "," & (tjlb_len2-tjlb_len1) & ")<>'00' and tjlb_sxh>=" & request("dwxh")' �����˳�����ǰ��
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-tjbb.asp?mode=101&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2)  

elseif mode=207 then
  'С��delete
  opendb()
  conn.execute "delete from tjlb where tjlb_code like'" + left(request("dq"),tjlb_len3)+"%'"'�������С��
  conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where left(tjlb_code," & tjlb_len2 & ")='" & left(request("dqcode2"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and tjlb_sxh>=" & request("dwxh")' �����˳�����ǰ��
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-tjbb.asp?mode=201&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") &"&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&gs_cc2=" & request("gs_cc2") & "&gs2=" & request("gs2") & "&page3=" & cpage3)

elseif mode=8 then
  'delete ��������/����
  opendb()
  if request("sort")="up" then'����
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh+1 where right(tjlb_code,"& (tjlb_len0-tjlb_len1) & ")='" & tjlb_str1 &"' and tjlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where tjlb_code='" + request("dq")+"'"
  else'����
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where right(tjlb_code,"& (tjlb_len0-tjlb_len1) & ")='" & tjlb_str1 &"' and tjlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh+1 where tjlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-tjbb.asp?mode=1&page1=" & cpage1)

elseif mode=108 then
  'delete ��������/����
  opendb()
  if request("sort")="up" then'����
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh+1 where left(tjlb_code," & tjlb_len1 & ")='" & left(request("dqcode1"),tjlb_len1) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len2) & ")='" & tjlb_str2 &"' and mid(tjlb_code,"& (tjlb_len1+1) & "," & (tjlb_len2-tjlb_len1) & ")<>'00' and tjlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where tjlb_code='" + request("dq")+"'"
  else'����
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where left(tjlb_code," & tjlb_len1 & ")='" & left(request("dqcode1"),tjlb_len1) &"' and right(tjlb_code,"& (tjlb_len0-tjlb_len2) & ")='" & tjlb_str2 &"' and mid(tjlb_code,"& (tjlb_len1+1) & "," & (tjlb_len2-tjlb_len1) & ")<>'00' and tjlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh+1 where tjlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-tjbb.asp?mode=101&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2)

elseif mode=208 then
  'delete С������/����
  opendb()
  if request("sort")="up" then'����
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh+1 where left(tjlb_code," & tjlb_len2 & ")='" & left(request("dqcode2"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and tjlb_sxh=" & (request("dwxh")*1-1) & ""
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where tjlb_code='" + request("dq")+"'"
  else'����
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh-1 where left(tjlb_code," & tjlb_len2 & ")='" & left(request("dqcode2"),tjlb_len2) &"' and mid(tjlb_code,"& (tjlb_len2+1) & "," & (tjlb_len3-tjlb_len2) & ")<>'00' and tjlb_sxh=" & (request("dwxh")*1+1) & ""
    conn.execute "update tjlb set tjlb_sxh=tjlb_sxh+1 where tjlb_code='" + request("dq")+"'"
  end if
  closedb()
  delaySecond(2)
  Response.Redirect ("marea-tjbb.asp?mode=201&page1=" & cpage1 & "&dqcode1="+request("dqcode1") & "&dqname1=" & request("dqname1") & "&page2=" & cpage2 & "&dqcode2="+request("dqcode2") & "&dqname2=" & request("dqname2") & "&gs_cc2=" & request("gs_cc2") & "&gs2=" & request("gs2") & "&page3=" & cpage3)

end if
%>    