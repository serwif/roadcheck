<%@ LANGUAGE="VBSCRIPT" %>
<%option explicit%>

<%
if request("register")=1 then
elseif session("username")="" or instr(session("power"),",0,")=0 then
    Response.Redirect "notlogin.asp?title=用户管理"
end if
%>

<!--#include file="fcommon.asp"-->
<!--#include file="dtp.asp"-->

<%
dim conn_system, mode,register, username, rs,rsMX,rsMX1,rsMX2,rsMX3, sql, errmsg, founderror, s, t, i, fl, memname,cpage,truename,password,workphone,handset,familialphone,FRMunit,FRMbusiness,workshj,FRMdw,FRMpcs
dim sday,FRMcsrq,FRMrjrq,FRMwhcd

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

if not isempty(request("sday")) then
    sday = request("sday")
else
    sday = date()
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
  <title>日志管理</title>
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
  
  function threadmenu(){
   location.href="mrizhi.asp?mode=1&FRMdw="+document.form.FRMdw.options[document.form.FRMdw.selectedIndex].value
  }
  
  function threadmenu1()
  {
	//alert(document.form.odq1.value);
	//alert("mrizhi.asp?mode=1&FRMdw="+document.form.odq1.value+"&FRMpcs="+document.form.FRMpcs.options[document.form.FRMpcs.selectedIndex].value);
    location.href="mrizhi.asp?mode=1&FRMdw="+document.form.odq1.value+"&FRMpcs="+document.form.FRMpcs.options[document.form.FRMpcs.selectedIndex].value;
   return false;
  }


  //-->

  function check()
{
   location.href="mrizhi.asp?mode=1&sday=" + document.all.afsj.value; 
   return false; 
}
  </script>

  <body>
  <%noRightClick()%>
  <table width="90%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr bgcolor=<%=skincolor()%> height="28"><td align="center">
      <b>日志管理</b>
    </td></tr>
  </table>
<%
end sub

sub showctail()
%>
  </body>
  </html>
<%
end sub

if mode=1 then
'显示
  if not isEmpty(request("page")) then
    cpage = clng(request("page"))
  else
    cpage = 1
  end if
  showchead()
  'Response.Write "<br>"
  opendb()
   
  set rs=server.createobject("adodb.recordset")
  set rsMX=server.createobject("adodb.recordset")
  %>
  <table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="#eeeeee" align=left><input type="text" name="afsj" size="10" maxlength="12" readonly  value='<%=sday%>' onchange="check()" >
      <A onclick="show_cele_date(change2,'','',afsj)"><IMG align=top border=0 height=25 name=change2 src="images\calendar.gif" width=26></A>
    </td>
    <%rs.open "select * from olog where shj like '"&year(sday) &"-"&month(sday)&"-"&day(sday)&"%'",conn_system, 1, 1
    if rs.recordcount <> 0 then
      rs.movefirst
      rs.CacheSize = 5
      rs.PageSize = 10
      if cpage>rs.pagecount then cpage=1
      rs.AbsolutePage = cpage%>
      <td valign="bottom" align="right">第<%=cstr(cpage)%>页/共<%=cstr(rs.PageCount)%>页，共<font color="blue"><strong><%=cstr(rs.RecordCount)%></strong></font>条记录</td>
      <td align="right">
        <%if cpage <> 1 then%>
          [<a href="mrizhi.asp?mode=1&sday=<%=sday%>&page=<%=cstr(cpage-1)%>">上一页</a>]
        <%end if%>
        <%if cpage <> rs.PageCount then%>
          [<a href="mrizhi.asp?mode=1&sday=<%=sday%>&page=<%=cstr(cpage+1)%>">下一页</a>]
        <%end if%>
        <%if rs.PageCount > 1 then%>
	  <select name="select2"  onchange="goto(this)">
            <%for i = 1 to rs.PageCount%>
              <%if i = cpage then%>
                <option selected value="mrizhi.asp?mode=1&sday=<%=sday%>&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%else%>
                <option value="mrizhi.asp?mode=1&sday=<%=sday%>&page=<%=cstr(i)%>">到第 <%=cstr(i)%> 页</option>
              <%end if%>
             <%next%>
          </select>
        <%end if%>
        </td>
        </tr>
        <tr><td colspan="6">
          <table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr bgcolor=<%=skincolor()%>>
	      <td width=60 align=center>时间</td>
              <td width=40 align=center>用户</td>
              <td width=150 align=center>操作描述</td>
              <td width=40 align=center>备注</td>
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
                <td><%=rs("shj")%></td>
                <td><%=rs("username")%></td>
                <td><%=rs("czms")%></td>
                <td>
                  <%if rs("bz")="DL" then
                    response.write("登录")
                  elseif rs("bz")="ZJ" then
                    response.write("增加案件")
                  elseif rs("bz")="XG" then
                    response.write("修改案件")
                  elseif rs("bz")="DL" then
                    response.write("删除案件")
                  else
                    response.write("&nbsp;")
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
  <!--<br><br>-->
    <!--<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr>-->
        <td valign="bottom" align="right"></td>  
        <td align="right">
        </td>
      </tr>
      <tr><td colspan="6"><hr size=1 width=100% noshade></td></tr>
      <tr><td align="center" colspan="6"><font size="6">没有记录</font></td></tr>
    </table>

  <%end if
  rs.close
  set rs=nothing
  closedb()
  showctail()




end if
%>    