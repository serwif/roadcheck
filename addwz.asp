<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if session("username")=""  or (instr(session("power"),",1,")=0 and instr(session("power"),",2,")=0) then
  Response.Redirect("notlogin.asp")
end if
%>

<!--#include file="fcommon.asp"-->
<!--#include file="dtp.asp"-->
<%
dim conn, mode, username, rs, sql,rs1,rsMX, errmsg, founderror, s, t, i, fl, dq,odq, cpage,kpbm,st
dim unit_code,unit_name,ajlb_code,fxlb_code,afsj,ajjs,czshj,explain,sday,ajbh,zbzcy,zp
dim wzlb_code,wzlb_name,bt,nr,zz

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
if not isempty(request("unit_code")) then
    unit_code = request("unit_code")
else
    unit_code =""
end if
if not isempty(request("bt")) then
    bt = request("bt")
else
    bt = ""
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
  <title>文章管理</title>
  <link rel="stylesheet" type="text/css" href="./main.css">
  </head>
<script language="javascript">
  <!--
  function surfto(list){
   var myindex1=list.selectedIndex;
   if (myindex1!=0 & myindex1!=1){ location.href=list.options[list.selectedIndex].value }
  }
  function goto(list){
   location.href=list.options[list.selectedIndex].value
  }
  //-->
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
   location.href="addwz.asp?mode=1&DisDate=8&sday=" + document.all.afsj.value+"&unit_code="+document.form1.unit_code.value; 
   return false;   
}

function check_form() 
{ 
  if(trim(document.form1.bt.value)==""){
   alert("请完整填入标题!"); 
   return false; 
   }
  if(trim(document.form1.zz.value)==""){
   alert("请完整填入作者!"); 
   return false; 
   }
  document.form1.nr.value=document.form1.doc_html.value;
  if(trim(document.form1.nr.value)==""){
   alert("请完整填入内容!"); 
   return false; 
   }
  return true;
} 

function loadForm()
{
  document.form1.doc_html.value=document.form1.nr.value;
  return true
}

</script>
  <%if mode=3 then'修改模式,窗体导入时将文章内容赋值给内容编辑器%>
    <body onload="loadForm()">
  <%else%>  
    <body>
  <%end if%>
  <%noRightClick()%>
  <!--<table width="90%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28"><td align="center">
      <b>文章管理</b>
    </td></tr>
  </table>-->
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
    'rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn, 1, 1%>
    <form name="form1" method="post" onsubmit="return check_form()" action="addwz.asp?mode=2">
  <%else
    opendb()
    set rs1=server.createobject("adodb.recordset")
    set rsMX=server.createobject("adodb.recordset")
    'rs1.open "select * from unit where right(unit_code,"& (unit_len0-unit_len1) & ")='" & unit_str1 &"' order by unit_sxh", conn, 1, 1
    set rs=server.createobject("adodb.recordset")
    rs.open "select * from wz where bh='" + request("odq") + "'", conn, 1, 1
    %>
    <form name="form1" method="post" onsubmit="return check_form()" action="addwz.asp?mode=3&odq=<%=request("odq")%>">
  <%end if%>
  <table width="530" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>文章</b></td>
    </tr>
    <tr>
      <td align=center>
        <table width=100%><tr>
        </table>
      </td>
    </tr>
    <tr>
      <td align="center"><br><font color=red>
        <%if Trim(ErrMsg) <> "" then%>
          <%=errmsg%>
        <%else%>
          请填写文章内容。
          <% if mode=2 then%>
            <input name="odq1" type="hidden" value="">
          <% else %>
            <input name="odq1" type="hidden" value="<%=rs("bh")%>">
          <% end if %>
        <%end if%>
        </font><br><br>
      </td>
    </tr>
    <tr>
      <td align=center>
        <table width="100%" border="0" cellspacing="1" bgcolor="#cccccc">
          <!--DWLayoutTable-->
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">文章类别：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF">
            <select name="lb" style="HEIGHT:17px;WIDTH:119px" >
            <%rsMX.open "select * from wzlb order by wzlb_sxh", conn, 1, 1
            while not rsMX.EOF
              %>
              <%if mode=2 then%>
                <option value="<%=trim(rsMX("wzlb_code"))%>"><%=trim(rsMX("wzlb_name"))%></option> 
              <%else%>
                <option value="<%=trim(rsMX("wzlb_code"))%>"<%if rs("lb")=rsMX("wzlb_code") then %> selected <% end if %>><%=trim(rsMX("wzlb_name"))%></option>
              <%end if%>
              <%rsMX.MoveNext 
            WEND
            rsMX.close%>
            </select>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">标题：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="bt" size="60" maxlength="120"  value=''>
              <%else%>
                <input type="text" name="bt" size="60" maxlength="120" value='<%=rs("bt")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">作者：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <input type="text" name="zz" size="20" maxlength="20"  value=''>
              <%else%>
                <input type="text" name="zz" size="20" maxlength="20" value='<%=rs("zz")%>'>
              <%end if%>
            </td>
          </tr>
          <tr>
            <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">文章内容：</td>
            <td height="23" colspan="7" bgcolor="#FFFFFF"> 
              <%if mode=2 then%>
                <object id=doc_html style="LEFT: 0px; TOP: 0px" data=editor/editor.html width=450 height=660 type=text/x-scriptlet VIEWASTEXT>
                </object> 
                <input type="hidden" name="nr" value="" > 
              <%else%>
                <object id=doc_html style="LEFT: 0px; TOP: 0px" data=editor/editor.html width=450 height=660 type=text/x-scriptlet VIEWASTEXT>
                </object> 
                <input type="hidden" name="nr" value='<%=replace(rs("nr"),"'","&quot;")%>'>
              <%end if%>
            </td>
          </tr>
        </table>
        <p> 
        <input class="buttonface" type="submit" name="Submit" value=" 提 交 ">
        &nbsp; 
        <INPUT class="buttonface" type=reset onclick="{if(confirm('该项操作要清除全部的内容，您确定要清除吗?')){return true;}return false;}" value=" 重 写 " id=reset1 name=reset1>
        </p>   
        </div>
      </td>
    </tr>
    </table>
  </form>
<%
  if mode = 2 then
    'rs1.close
    set rs=nothing
    closedb()
  elseif mode = 3  then
    'rs1.close
    rs.close
    set rs=nothing
    closedb()
  end if
  showctail
end sub

sub ShowInputForm3(ErrMsg)
  'on error resume next
  showchead()%>
  <form method="POST" action="addwz.asp?mode=5&username=<%=username%>" name="input3">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td align="right">
        [<a href="addwz.asp?mode=1">返回列表</a>]
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
  if right(session("unit_code"),unit_len0-unit_len2) = unit_str2 and mid(session("unit_code"),unit_len1+1,unit_len2-unit_len1)="00" then
    sql="select * from wz where unit_code like '"+session("unit_code")+"%' "
  else
    sql="select * from wz where unit_code='"+session("unit_code")+"' "
  end if
  sql=sql+" order by shj"
  'response.write sql
  rs.open sql, conn, 1, 1%>
  <form name="form1">
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td height="23"  align="left" bgcolor="#FFFFFF" >
        <input name="unit_code" type="hidden" value="<%=request("unit_code")%>">
      </td>
    </tr>
  </tabel>
  </form>
  <%if rs.recordcount <> 0 then
  rs.movefirst
  rs.CacheSize = 5
  rs.PageSize = 10
  if cpage>rs.pagecount then cpage=1
  rs.AbsolutePage = cpage%>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
    </tr>
    <tr>
      <td valign="bottom">第<%=cstr(cpage)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>个花名册记录</td>
      <td align="right">
        [<a href="addwz.asp?mode=5">查找</a>]
        <%if instr(session("power"),",1,")=1 then %>
          [<a href="addwz.asp?mode=2&unit_code=<%=session("unit_code")%>">添加</a>]
        <%end if %>
        <%if cpage <> 1 then%>
          [<a href="addwz.asp?mode=1&page=<%=cstr(cpage-1)%>">上一页</a>]
        <%end if%>
        <%if cpage <> rs.PageCount then%>
          [<a href="addwz.asp?mode=1&page=<%=cstr(cpage+1)%>">下一页</a>]
        <%end if%>
        <%if rs.PageCount > 1 then%>
          <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage then%>
                <option selected value="addwz.asp?mode=1&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="addwz.asp?mode=1&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
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
            <td width=10% align=center>类别</td>
            <td width=10% align=center>标题</td>
            <td width=5% align=center>作者</td>
            <!--<td width=30% align=center>文章内容</td>-->
            <td width=10% align=center>时间</td>
            <%if instr(session("power"),",2,")>0 then %>
              <td width=15% align=center>操作</td>
            <%end if%>
          </tr>
          <%fl = False
          for i = 1 to rs.PageSize
            if not rs.EOF then
              if fl then%>
                <tr bgcolor="#eeeeee">
              <%else%>
                <tr>
              <%end if%>
              <td align=center>
                <%
                rs1.open "select * from wzlb where wzlb_code='" & rs("lb") & "'",conn,1,1
                if rs1.recordcount=0 then
                  response.write "&nbsp;"
                else
                  response.write rs1("wzlb_name")
                end if
                rs1.close
                %>
              </td>
              <td align=center><%=rs("bt")%></td> 
              <td align=center><%=rs("zz")%></td> 
              <!--<td align=center>
                <%if rs("nr")="" then 
                  response.write "&nbsp;"
                else
                  if len(rs("nr"))>50 then
                    response.write left(rs("nr"),50) & "...."
                  else
                    response.write rs("nr")
                  end if
                end if%>
              </td> -->
              <td align=center>
                <%if len(rs("shj"))=14 then
                  response.write todatestr(left(rs("shj"),8)) & " " & totimestr(right(rs("shj"),6))
                else
                  response.write "&nbsp;"
                end if
                %>
              </td> 
              <%if instr(session("power"),",2,")>0 then%>
                <td align=center>
                  <a href="addwz.asp?mode=3&unit_code=<%=request("unit_code")%>&odq=<%=trim(rs("bh"))%>"><img src="./images/edit.gif" border=0></a>
                  <a href="addwz.asp?mode=4&unit_code=<%=request("unit_code")%>&dq=<%=trim(rs("bh"))%>&bt=<%=trim(rs("bt"))%>"><img src="./images/del.gif" border=0></a>
                </td>
              <%end if%>
            </tr>
            <%rs.MoveNext
            fl = not fl
          end if
        next%>
      </table>
    </td>
  </tr>
</table>
  <%else%>
    <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>
        <td align="right">
          <%if instr(session("power"),",1,")=1 then%>
          [<a href="addwz.asp?mode=2&unit_code=<%=session("unit_code")%>">添加</a>]
          <%end if %>
        </td>
      </tr>
      <tr><td><hr noshade size=1 width=100%></td></tr>
      <tr><td align="center"><font size="6">没有文章记录</font></td></tr>
    </table>
  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()
elseif mode=2 or mode=3 then
  '添加及修改
  if trim(request("bt")) <> "" then
    if trim(request("unit_code2"))="" then
      unit_code=request("unit_code1")
    else
      unit_code=request("unit_code2")
    end if
    if unit_code="" then
      unit_code=session("unit_code")
    end if
    username=trim(session("username"))
    czshj=datetostr(now()) &timetostr(now) & "00"
    FoundError=false
    ErrMsg=""
    if not FoundError then
      if mode=2 then
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册，判断有无重复的同一时间做两件事
        rs.open "select bh from wz where unit_code='"+unit_code+"' and bt='" + request("bt") + "'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "标题重复"
          FoundError = True
        end if
        rs.close
        set rs=nothing
        closedb()
      else
        '看改过的案件编号是否存在
        opendb()
        set rs=server.createobject("adodb.recordset")
        '查找是否有重复的注册，判断有无重复的同一时间做两件事
        rs.open "select bh from wz where unit_code='"+unit_code+"' and bt='" + request("bt") + "' and bh<>'"&request("odq") &"'", conn, 1, 1
        if rs.recordcount<>0 then
          if ErrMsg <> "" then ErrMsg = ErrMsg + "<br>"
          ErrMsg = ErrMsg + "姓名重复"
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
      if mode=2 then
	'判断是否已经存在同一时间段同一个人的工作记录
	opendb()
	set rs=server.createobject("adodb.recordset")
        if FoundError=true then
	  set rs=nothing
	  closedb()
	  ShowInputForm1 mode,errmsg
	else
          kpbm=right(cstr(year(date)),4)
	  if month(date)<10 then
	    kpbm=kpbm&"0"&cstr(month(date))
	  else
	    kpbm=kpbm&cstr(month(date))
	  end if
	  if day(date)<10 then
	    kpbm=kpbm&"0"&cstr(day(date))
	  else
	    kpbm=kpbm&cstr(day(date))
	  end if
	  rs.open "select bh from wz where bh like'" & kpbm & "%' order by bh desc", conn,1,1
	  if rs.recordcount=0 then
	    kpbm=kpbm&"0001"
	  else
	    rs.movefirst
	    st=cstr(cint(right(rs("bh"),4))+1)
	    for i=len(st) to 3
	      st="0"&st
	    next
            kpbm=kpbm&st
	  end if
          rs.close
          'response.write request("nr")
          rs.open "wz",conn,1,3
          rs.addnew
          rs("bh")=kpbm
          rs("bt")=request("bt")
          rs("zz")=request("zz")
          rs("unit_code")=unit_code
          rs("nr")=request("nr")
          rs("lb")=request("lb")
          rs("username")=username
          rs("shj")=czshj
          rs.update
          rs.close
          '保存增加
          conn.execute("insert into olog (shj,username,czms,bz) values ('"&now()&"','"&username&"','增加文章："&unit_name&","&request("bt")&"','ZJWZ')")
	  set rs=nothing
	  closedb()
	  Response.Redirect "addwz.asp?mode=1"
	end if
      else
	opendb()
        'response.write request("nr")
        'set rs=server.createobject("adodb.recordset")
        'response.write "update wz set lb='"&request("lb")&"',bt='"&request("bt")&"',unit_code='"&unit_code&"',zz='"&request("zz")&"',nr='"&request("nr")&"' where bh='"&request("odq") &"'"
        conn.execute("update wz set lb='"&request("lb")&"',bt='"&request("bt")&"',unit_code='"&unit_code&"',zz='"&request("zz")&"',nr='"&request("nr")&"' where bh='"&request("odq") &"'")
        conn.execute("insert into olog (shj,username,czms,bz) values ('"&now()&"','"&username&"','修改花名册："&unit_name&","&request("bt")&"','XGWZ')")
	closedb()
	Response.Redirect "addwz.asp?mode=1"
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
        [<a href="addwz.asp?mode=1">返回</a>]
     </td>
    </tr>
    <tr><td><hr size="1" noshade width=100%></td></tr>
    <tr><td align="center">
      <br><br>
      真的要删除这个文章“<%=request("bt")%>”？
      <br><br>
      [<a href="addwz.asp?mode=7&dq=<%=request("dq")%>&bt=<%=request("bt")%>">是的</a>]
      &nbsp;&nbsp;&nbsp;[<a href="addwz.asp?mode=1">算了</a>]
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
    set rs1=server.createobject("adodb.recordset")
    sql=""
    if trim(request("dq")) <> "" then
      sql="(nr like '%" + trim(request("dq")) + "%')"
    end if
    rs.open "select * from wz where " + sql, conn, 1, 1
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
            [<a href="addwz.asp?mode=1">返回列表</a>] 
            [<a href="addwz.asp?mode=5">继续查找</a>]
         </td>
        </tr>
        <tr><td>
          <%rs.movefirst
    kpbm=request("odq")
    if  kpbm= "" then
      kpbm = rs("bh")
    else
      rs.Find "bh= '" + kpbm +"'"
      if rs.EOF then
        rs.movefirst
        kpbm = rs("bh")
      end if
    end if
    %>
  <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td>
        [共<strong><font color="blue"><%=rs.recordcount%></font></strong>条记录 </a>]
        <%
        rs.MovePrevious
        if not rs.BOF then%>
          <a href="addwz.asp?mode=5&dq=<%=request("dq")%>&odq=<%=trim(rs("bh"))%>">[上一页]</a>
        <%end if
        rs.Move 2
        if not rs.EOF then%>
          <a href="addwz.asp?mode=5&dq=<%=request("dq")%>&odq=<%=trim(rs("bh"))%>">[下一页]</a>
        <%end if
        rs.MovePrevious
        %>
      </td>
    </tr>
    <tr bgcolor=<%=skincolor()%> height="28">
      <td align="center"><b>文章</b></td>
    </tr>
    <tr>
      <td align=center>
        <table width="100%" border="0" cellspacing="1" bgcolor="#cccccc">
          <!--DWLayoutTable-->
            <tr>
              <td height="23" colspan="1" bgcolor="#eeeeee" align="right">文章类别：</td>
              <td height="23" colspan="7" bgcolor="#FFFFFF">
                <%
                rs1.open "select * from wzlb where wzlb_code='" & rs("lb") & "'",conn,1,1
                if rs1.recordcount=0 then
                  response.write "&nbsp;"
                else
                  response.write rs1("wzlb_name")
                end if
                rs1.close
                %>
              </td>
            </tr>
            <tr>
              <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">标题：</td>
              <td height="23" colspan="7" bgcolor="#FFFFFF">
                <%response.write rs("bt")%>
              </td>
            </tr>
            <tr>
              <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">作者：</td>
              <td height="23" colspan="7" bgcolor="#FFFFFF">
                <%response.write rs("zz")%>
              </td>
            </tr>
            <tr>
              <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">文章内容：</td>
              <td height="23" colspan="7" width=400 bgcolor="#FFFFFF">
                <%response.write replace(rs("nr"),request("dq"),"<font color=red>"&request("dq")&"</font>")%>
              </td>
            </tr>
            <tr>
              <td height="23" colspan="1" bgcolor="#eeeeee"  align="right">发布时间：</td>
              <td height="23" colspan="7" width=400 bgcolor="#FFFFFF">
                <%if len(rs("shj"))=14 then
                  response.write todatestr(left(rs("shj"),8)) & " " & totimestr(right(rs("shj"),6))
                else
                  response.write "&nbsp;"
                end if
                %>
              </td>
            </tr>
            </table>
          </div>
        </td>
      </tr>
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

elseif mode=7 then
  'delete
  opendb()
  conn.execute "delete from wz where bh='" + request("dq")+"'"
  '保存删除案件日志
  conn.execute("insert into olog (shj,username,czms,bz) values ('"&now()&"','"&username&"','删除文章："&request("bt")&"','SCWZ')")
  closedb()
  delaySecond(2)
  Response.Redirect ("addwz.asp?mode=1&unit_code=" & request("unit_code"))
elseif mode=102 then
  ShowInputForm1 2,""
elseif mode=103 then
  ShowInputForm1 3,""
end if
%>    