<!--#include file="conn.asp" -->
<!--#include file="const.asp"-->
<!--#include file="function.asp" -->
<%
if cint(request("model"))=1 then
	jsnavigate(1)
else
	jsnavigate(2)
end if
set rs=nothing
conn.close
set conn=nothing
	
Function jsnavigate(strtr)
	dim javastr,i	
	javastr="<table width=98% border=0 align=center cellspacing=0 cellpadding=0 bgcolor="""&CenterBColor&""">"
	javastr=javastr+TTitle("center_1","À¸Ä¿µ¼º½")	
	javastr=javastr+"<tr>"
	javastr=javastr+"<td bgcolor="""&CenterCColor&""" background="""&CenterCImg&""" align=right valign=top>"
	javastr=javastr+"<table width=98% border=0 align=right cellspacing=2 cellpadding=0 style=""TABLE-LAYOUT: fixed"">"
	i=1
	sql="select BigClassName,BigClassType,BigTemplate,BigClassView from BigClass order by BigClassID"
	rs.open sql,conn,1,1
	set rs1=server.createobject("adodb.recordset")
	while not rs.EOF
'		if rs(3)<>0 then
			if strtr=1 then
				javastr=javastr+"<tr>"
			else
				if i mod 2=1 then javastr=javastr+"<tr width=50% >"
			end if
				if rs(1)=0 then
					BigClassType="&BigClassType=0"
				else
					BigClassType=""
				end if
				if rs(2)>1 then
					BigTemplate="&BigTemplate="&rs(2)
				else
					BigTemplate=""
				end if		
			javastr=javastr+"<td valign=top><b><a Class=BigClass href=""BigClass.asp?BigClassName="&rs(0)&BigClassType&BigTemplate&""">"&rs(0)&"</b><br>"
			if strtr=1 then javastr=javastr+"</td><td width=80% >"		
			sql="select SmallClassName,SmallClassType from SmallClass where BigClassName='"&rs(0)&"' order by SmallClassID"
			rs1.open sql,conn,1,1
			while not rs1.EOF
				if rs1(1)=0 then
					SmallClassType="&SmallClassType=0"
				else
					SmallClassType=""
				end if
				if rs(2)>1 then
					BigTemplate="&BigTemplate="&rs(2)
				else
					BigTemplate=""
				end if		
				javastr=javastr+"<a Class=LeftMenu href=""SmallClass.asp?BigClassName="&rs(0)&BigTemplate&"&SmallClassName="&rs1(0)&SmallClassType&""">"&rs1(0)&"</a>¡¡"		
				rs1.MoveNext
			wend
			rs1.close		
			javastr=javastr+"</td>"
			if strtr=1 then
				javastr=javastr+"</tr>"
			else		
				if i mod 2=0 then javastr=javastr+"</tr>"
			end if
			i=i+1
'		end if
		rs.MoveNext
	wend
	rs.Close
	set rs1=nothing	
	javastr=javastr+"</table></td></tr>"
	javastr=javastr+InTable("middle1")
	javastr=javastr+"</table>"
	response.write ("document.write('"&javastr&"')")
	response.end
End Function
%>