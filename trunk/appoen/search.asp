<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function_title.asp" -->
<%
dim count,onecount,br
if request("align")<>"w" then br="<br>"
rs.open "select BigClassName,SmallClassName from SmallClass order by SmallClassID DESC",conn,1,1
%>
	var onecount;
	onecount=0;
	subcat = new Array();
	<%
	count = 0		
	do while not rs.eof
		%>
		subcat[<%=count%>] = new Array("<%=rs(1)%>","<%=rs(0)%>");
		<%
		count = count + 1
		rs.movenext
	loop
	rs.close
	%>
	onecount=<%=count%>;
	function changelocation(locationid)
	{document.myform.SmallClassName.length = 0;
	var locationid=locationid;
	var i;
	for (i=0;i < onecount; i++)
	{if (subcat[i][1] == locationid)
	{document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
	}}}
<%
javastr="<table border=""0"" width=""100%"" align=center cellspacing=""0"" cellpadding=""0"" bgcolor="""&LeftBColor&""">"
javastr=javastr+TTitle("left","资料搜索")
javastr=javastr+"<form method=""post"" name=""myform"" action=""showsearch.asp"">"
javastr=javastr+"<tr><td width=""100%"" bgcolor="""&LeftCColor&"""  background="""&LeftCImg&""" align=center valign=""middle""><br>"
response.write ("document.write('"&javastr&"');")

javastr="<select name=""action"" size=""1"">"
javastr=javastr+"<option value="""">不指定条件</option>"
javastr=javastr+"<option value=""title"" "
if request("action")="title" then javastr=javastr+"selected"
javastr=javastr+">按标题</option>"
javastr=javastr+"<option value=""content"" "
if request("action")="content" then javastr=javastr+"selected"
javastr=javastr+">按内容</option>"
javastr=javastr+"<option value=""author"" "
if request("action")="author" then javastr=javastr+"selected"
javastr=javastr+">按作者</option>"
javastr=javastr+"<option value=""original"" "
if request("action")="original" then javastr=javastr+"selected"
javastr=javastr+">按来源文号</option>"
javastr=javastr+"</select>"&br
response.write ("document.write('"&javastr&"');")

javastr="<select name=""BigClassName"" onChange=""changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)"" size=""1"">"
javastr=javastr+"<option selected value="""">不指定大类</option>"
set rs=conn.execute("select BigClassName from BigClass order by BigClassID")
do while not rs.eof
javastr=javastr+"<option value="""&rs(0)&""">"&rs(0)&"</option>"
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
javastr=javastr+"</select>"&br
response.write ("document.write('"&javastr&"');")


javastr="<select name=""SmallClassName"">"
javastr=javastr+"<option selected value="""">不指定小类</option>"
javastr=javastr+"</select>"&br
javastr=javastr+"<input type=""text"" name=""key"" size=10 value="""&request("key")&""" maxlength=""50"">"
javastr=javastr+"&nbsp;<input type=""submit"" name=""Submit"" value=""模"" class=""submit"">"
javastr=javastr+"&nbsp;<input type=""submit"" name=""Submit"" value=""精"" class=""submit""><br>"
javastr=javastr+space(10)
javastr=javastr+"</td></tr>"
javastr=javastr+"</form>"
javastr=javastr+InTable("bottoml")
javastr=javastr+"</table>"
response.write ("document.write('"&javastr&"');")
response.end
%>