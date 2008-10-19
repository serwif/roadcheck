<!--#include file="conn.asp" -->
<!--#include file="config.asp" -->
<!--#include file="user/user_config.asp"-->
<%
dim newsID,titleurl
NewsID=cint(request("ID"))

sql="select fname,BigClassName,SmallClassName,titleurl from News where checked="&true&" and NewsID=" & NewsID
rs.open sql,conn,1,1
if rs.eof then
%>
	<script language=javascript>
	history.back()
	alert("你要查看的内容不存在或已经给管理员锁定，请与管理联系！")
	</script>
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing	
	Response.End 
else	
	fname=rs(0)
	BigClassName=rs(1)
	SmallClassName=rs(2)
	titleurl=rs(3)
	rs.Close
end if


'-----------检查会员
if ShowUserLogin=1 then
	dim ReaderLevel,UserLevel
	sql="select ReaderLevel from SmallClass where BigClassName='"&BigClassName&"' and SmallClassName='"&SmallClassName&"'"
	rs.open sql,conn,1,1
	ReaderLevel=rs(0) 
	rs.close
	
 	if ReaderLevel<>0 then	'如果有等级
      '  if readfree=0 then  '如果只允许会员浏览
		   if (isnull(session("xpUser")) or isnull(session("xpPwd")) or session("xpUser")="" or session("xpPwd")="") then
				  set rs=nothing
				  conn.close 
				  set conn=nothing
				  response.write"你只有注册成为APPOEN.COM用户并且登陆才能够查看该内容，请先登陆或者注册!"
				  %>
					     <script language=javascript>
					   window.open('user/userlogin.asp',	'userlogin','width=800,height=400','status=on','location=on','toolbar=on','scrollbars=on')
					   </script>
					  <%
	
				 ' Response.Redirect "user/userlogin.asp"
				  Response.End 
		   end if
			
		   if not(isnull(session("xpUser")) or isnull(session("xpPwd")) or session("xpUser")="" or session("xpPwd")="") then
				sql="select UserLevel,lockuser,LimitPoint,readnews,UserPoint from Users where UserName='"&Session("xpUser")&"' and PassWord='"&Session("xpPwd")&"'"
			  rs.Open sql,conn,1,3
			  if rs.eof then
					   rs.close
					   set rs=nothing
					   conn.close
					   set conn=nothing
					    %>
					     <script language=javascript>
					   history.back()
					   alert("会员专栏，请先登陆。")
					   </script>
					  <%
					   Response.End 
			  else
				 dim rs1,theLimitPoint
			     set rs1=conn.execute("SELECT LimitPoint FROM UserGrade where id="&rs(0)&"")
				 theLimitPoint=rs1(0)
				 rs1.close
				  set rs1=nothing	
				   if rs(0)<7 then
					  if rs(0)<ReaderLevel then
						      rs.close
						       set rs=nothing
						       conn.close
						       set conn=nothing					
						        %>
						        <script language=javascript>
						        history.back()
						         alert("你的等级达不到本栏目所要求的级别第<%=ReaderLevel%>级，无权查看！请与管理员联系！。")
						         </script>
						       <%
						       Response.End 
					   elseif rs(1)=true then
						        rs.close
						         set rs=nothing
						         conn.close
						         set conn=nothing					
						          %>
						         <script language=javascript>
						          history.back()
						          alert("你的帐号被锁定，请联系管理员。")
						          </script>
						          <%
						           Response.End 
					   elseif rs(2)>=theLimitPoint then
						         rs.close
						         set rs=nothing
						         conn.close
						         set conn=nothing					
						         %>
						         <script language=javascript>
						          history.back()
						          alert("你的阅读次数已超过限制<%=theLimitPoint%>次，请与管理员联系！")
						          </script>
						          <%
						          Response.End 				
					    else
						          rs(3)=rs(3)+1
					        	   rs(4)=rs(4)+1
					         	   rs(2)=rs(2)+1
					        	   if  int(rs(4))=int(point(rs(0)+1)) then rs(0)=rs(0)+1 'end if
					          	   rs.Update
								   appoen_read
					    end if
					 else
                             appoen_read
					 end if
			
			 
			 end if

					  
				      rs.close
           end if
	   'end if
  else
     appoen_read
  end if
else
  appoen_read
end if
'-----------

sub appoen_read
conn.execute("update News Set Click=click+1 where NewsID=" & NewsID )

if titleurl="" or isnull(titleurl) then
	if fname<>"" or not isnull(fname) then
		set rs=nothing
		conn.close
		set conn=nothing	
		response.redirect "html/"&mid(fname,1,4)&"/"&mid(fname,5,2)&"/"&fname&"-1.htm"
	else
		set rs=nothing
		conn.close
		set conn=nothing	
		response.redirect "oldnews_html.asp?newsid="&newsid		
	end if
else
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect titleurl
end if
end sub
%>