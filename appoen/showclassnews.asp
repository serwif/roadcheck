<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<!--#include file="function.asp" -->
		 <%
	dim javastr,n
	reDim ArrayBigClassView(50),ArrayBigClassType(50),ArrayBigClassName(50),ArrayBigTempLate(50)
	dim RSCount
	dim totalNews,totalontop,classurl,msql,Thissql
	sql="select BigClassView,BigClassType,BigClassName,BigTempLate from BigClass order by BigClassID"
	rs.open sql,conn,1,1
	RSCount=rs.RecordCount
	i=1
	while not rs.EOF						
			ArrayBigClassView(i)=rs(0)
			arrayBigClassType(i)=rs(1)
			ArrayBigClassName(i)=rs(2)
			ArrayBigTempLate(i)=rs(3)
		i=i+1		
		rs.MoveNext
	wend
	rs.close
	
select case cint(request("fentitle"))
case 2
	jsbigclass2	
case 3
	jsbigclass3
case else
	jsbigclass1
end select

set rs=nothing
conn.close
set conn=nothing

Function jsbigclass1	
	javastr=""
	if RSCount>0 then	
		for i=1 to RSCount
			if ArrayBigClassType(i)=1 and ArrayBigClassView(i)=1 then
				BigClassName=ArrayBigClassName(i)
				sql="select newsid from News where BigClassName='" & BigClassName &"' and checked="&true&""
				rs.open sql,conn,1,1
				totalNews=rs.recordcount
				rs.close				
				sql="select newsid from News where BigClassName='" & BigClassName &"' and OnTop="&true&" and checked="&true&""
				rs.open sql,conn,1,1
				totalontop=rs.recordcount
				rs.close
				n=0
				javastr=javastr+showclassnews(BigClassName,BigTemplate,"center_2",1,0,totalNews,totalontop)
				javastr=javastr+"<br>"
			end if
		next		
		response.write ("document.write('"&javastr&"')")
	end if
End Function

Function jsbigclass2	
	if RSCount>0 then
		dim tr
		tr=0
		response.write ("document.write('<table cellspacing=0 cellpadding=0 width=98% border=0 align=center>');")
		for i=1 to RSCount
			BigClassName=ArrayBigClassName(i)
			if ArrayBigClassType(i)=1 and ArrayBigClassView(i)=1 then
				dim totalNews
				sql="select newsid from News where BigClassName='" & BigClassName &"' and checked="&true&""
				rs.open sql,conn,1,1
				totalNews=rs.recordcount
				rs.close			
				tr=tr+1	
				sql="select newsid from News where BigClassName='" & BigClassName &"' and OnTop="&true&" and checked="&true&""
				rs.open sql,conn,1,1
				totalontop=rs.recordcount
				rs.close
				if ArrayBigTemplate(i)>1 then BigTemplate="&BigTemplate"&ArrayBigTemplate(i)
				if tr mod 2=1 then response.write ("document.write('<tr>');")
				response.write ("document.write('<td valign=top width=50% >');"&vbcrlf)				
				response.write ("document.write('"&showclassnews(BigClassName,BigTemplate,"center_2",2,0,totalNews,totalontop)&"')"&vbcrlf)
				javastr="<br></td>"	
				if tr mod 2=0  then 
					javastr=javastr+"</tr>"
				else
					if i=RSCount then javastr=javastr+"</tr>"
				end if
				response.write ("document.write('"&javastr&"')"&vbcrlf)	
			'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////			
			if tr=2 then response.write ("document.write('<tr><td colspan=2 valign=middle align=center>"&jsad("首页大类1排下")&"</td></tr>');"&vbcrlf)
			if tr=4 then response.write ("document.write('<tr><td colspan=2 valign=middle align=center>"&jsad("首页大类2排下")&"</td></tr>');"&vbcrlf)
			if tr=6 then response.write ("document.write('<tr><td colspan=2 valign=middle align=center>"&jsad("首页大类3排下")&"</td></tr>');"&vbcrlf)
			if tr=8 then response.write ("document.write('<tr><td colspan=2 valign=middle align=center>"&jsad("首页大类4排下")&"</td></tr>');"&vbcrlf)
			'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
			end if
		next
		javastr="</table>"	
	else
		javastr=javastr+"<center><b>尚　无　大　类</b></center>"
	end if		
	response.write ("document.write('"&javastr&"')")
End Function

Function jsbigclass3	
	if RSCount>0 then
	dim model,picid	
	tr=0
	response.write ("document.write('<table cellspacing=0 cellpadding=0 width=100% border=0 align=center>');")	
	for i=1 to RSCount
		BigClassName=ArrayBigClassName(i)
		if ArrayBigClassType(i)=1 and ArrayBigClassView(i)=1 then
				dim totalNews
				sql="select newsid from News where BigClassName='" & BigClassName &"' and checked="&true&""
				rs.open sql,conn,1,1
				totalNews=rs.recordcount
				rs.close		
			tr=tr+1			
			sql="select top 1 newsid,model from News where BigClassName='" & BigClassName &"' and checked="&true&" and image>0 order by updatetime desc"
			rs.open sql,conn,1,1
			if not rs.eof then
				model=rs("model")
				picid=rs("newsid")
			else
				picid=0
			end if
			rs.close			
			sql="select newsid from News where BigClassName='" & BigClassName &"' and OnTop="&true&" and checked="&true&""
			rs.open sql,conn,1,1
			totalontop=rs.recordcount
			rs.close
			if tr mod 2=1 then response.write ("document.write('<tr>');"&vbcrlf)			
			response.write ("document.write('<td width=50% align=right valign=top>');"&vbcrlf)
			response.write ("document.write('"&showclassnews(BigClassName,BigTemplate,"center_2",3,picid,totalNews,totalontop)&"');"&vbcrlf)
			javastr="<br></td>"
			if tr mod 2=0 then
				javastr=javastr+"</tr>"
			else
				if i=RSCount then javastr=javastr+"</tr>"
			end if
			response.write ("document.write('"&javastr&"');"&vbcrlf)
			'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////			
			if tr=2 then response.write ("document.write('<tr><td colspan=2 valign=middle align=center>"&jsad("首页大类一排下")&"</td></tr>');"&vbcrlf)
			if tr=4 then response.write ("document.write('<tr><td colspan=2 valign=middle align=center>"&jsad("首页大类二排下")&"</td></tr>');"&vbcrlf)
			if tr=6 then response.write ("document.write('<tr><td colspan=2 valign=middle align=center>"&jsad("首页大类三排下")&"</td></tr>');"&vbcrlf)
			if tr=8 then response.write ("document.write('<tr><td colspan=2 valign=middle align=center>"&jsad("首页大类四排下")&"</td></tr>');"&vbcrlf)
			'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
		end if		
	next
	javastr="</table>"
else
	javastr=javastr+"<center><b>尚　无　大　类</b></center>"
end if
	response.write ("document.write('"&javastr&"');")
End Function

'===========================================================================================================
Function ShowClassNews(strBigClassName,strtemplate,strtitlemodel,strmodel,strpicid,strtotal,strontop)
if strtemplate>1 then
	BigTemplate="&BigTemplate="&strtemplate
else
	BigTemplate=""
end if
classurl="BigClass.asp?BigClassName="&strBigClassName&BigTemplate
msql=NoContent&" from News where BigClassName='" & strBigClassName &"' and checked="&true&" "
			ShowClassNews="<table cellspacing=0 cellpadding=0 border=0 align=center width=95% style=""BORDER-LEFT: "&CenterTColor&" 1px double; BORDER-RIGHT: "&CenterTColor&" 1px double;BORDER-BOTTOM: "&CenterTColor&" 1px double;"" bgcolor="&centerccolor&" background="""&CenterCImg&""">"
			totalNews=strtotal
			ShowClassNews=ShowClassNews+TTitle(strtitlemodel,strBigClassName)			
			ShowClassNews=ShowClassNews+"<tr><td align=middle valign=top height=100><table width=100% border=0 align=right cellpadding=0 cellspacing=4><tr>"
			if showClassImg=1 and strpicid<>0 then
				ShowClassNews=ShowClassNews+"<td align=middle valign=top height=100 width=18% >"
				ShowClassNews=ShowClassNews+imagefile(strpicid,1,40,50)
				ShowClassNews=ShowClassNews+"</td>"
			end if
			ShowClassNews=ShowClassNews+"<td valign=top style=""WORD-WRAP: break-word"" class=MainContentS>"
			if strontop=0 then
				Thissql="select top " & MaxNewsList & msql & " order by updatetime DESC"
				showontop=""
				ShowClassNews=ShowClassNews+ClassTitle(Thissql,strpicid,strmodel,showontop)
			elseif totalontop>=MaxNewsList then
				Thissql="select top " & MaxNewsList & msql & " and OnTop="&true&" order by updatetime DESC"
				showontop="↑"				
				ShowClassNews=ShowClassNews+ClassTitle(Thissql,strpicid,strmodel,showontop)			
			elseif strontop>0 and strontop<MaxNewsList then
				Thissql="select top " & totalontop & msql & " and OnTop="&true&" order by updatetime DESC"
				showontop="↑"
				ShowClassNews=ShowClassNews+ClassTitle(Thissql,strpicid,strmodel,showontop)
				Thissql="select top " & MaxNewsList-totalontop & msql & " and OnTop=0 order by updatetime DESC"
				showontop=""				
				ShowClassNews=ShowClassNews+ClassTitle(Thissql,strpicid,strmodel,showontop)
			end if							
			ShowClassNews=ShowClassNews+"</td></tr></table></td></tr></table>"
End Function

Function ClassTitle(strsql,strpicid,strmodel,strontop)
ClassTitle=""
if strmodel=1 then
	rs.open strsql,conn,1,1
	while not rs.EOF
		n=n+1
		if cint(request("fentitle"))=1 then
		if n mod 2=1 then ClassTitle=ClassTitle+"<tr>"
		else
		ClassTitle=ClassTitle+"<tr>"
		end if
		ClassTitle=ClassTitle+"<td style=""WORD-WRAP: break-word"">"
		ClassTitle=ClassTitle+shownewf
		ClassTitle=ClassTitle+showTitle("MainContentS",20)&showImg
		ClassTitle=ClassTitle & showOntop 
		ClassTitle=ClassTitle+showclick & showOntop 
		'ClassTitle=ClassTitle+shownew
		ClassTitle=ClassTitle+"</td>"
		if cint(request("fentitle"))=1 then
		if n mod 2=0 then ClassTitle=ClassTitle+"</tr>"
		else
		ClassTitle=ClassTitle+"</tr>"
		end if
		rs.MoveNext
	wend		
	rs.close
elseif strmodel=2 then
	rs.open strsql,conn,1,1
	while not rs.EOF
		ClassTitle=ClassTitle+Shownewf
		ClassTitle=ClassTitle+ShowTitle("MainContentS",28)
		'ClassTitle=ClassTitle+ShowImg
		'ClassTitle=ClassTitle+ShowTime
		'ClassTitle=ClassTitle+showclick & showOntop
		ClassTitle=ClassTitle+"<br>"		
		rs.MoveNext
	wend
	rs.close
else
	dim maxlen
	rs.open strsql,conn,1,1
	while not rs.EOF
		maxlen=34
		if showClassImg=1 and strpicid<>"" then maxlen=28	
		if showOntop<>"" then maxlen=maxlen-2
		if rs("image")>0 then maxlen=maxlen-4
		if showClassImg=0 or strpicid="" then ClassTitle=ClassTitle+"&nbsp;"
		ClassTitle=ClassTitle+Shownewf		
		ClassTitle=ClassTitle+ShowTitle("MainContentS",maxlen)
		ClassTitle=ClassTitle+ShowImg
		'ClassTitle=ClassTitle+ShowTime
		ClassTitle=ClassTitle+Showclick
		ClassTitle=ClassTitle+showOntop
		ClassTitle=ClassTitle+"<br>"		
		rs.MoveNext
	wend
	rs.close
end if
end Function

function jsad(stradd)	'大类中调用广告，stradd为调用地址
dim add,mixid,maxid
if stradd<>"" then
	add="AdAdd='"&stradd&"' and "
else
	add="AdAdd='' and "
end if
sql="Select * from ad where "&add&" checked="&true&""
rs.Open sql, conn, 1, 3
if rs.bof or rs.eof then
	jsad=""
	rs.close
else
	rs.movefirst
	minid=rs("id")
	rs.movelast
	maxid=rs("id")
	rs.close	
	randomize
	x=fix((maxid-minid+1)*rnd)+minid	
	sql = "Select * From ad where "&add&" id="&x&" and checked="&true&""
	rs.Open sql,conn, 1,3	
	if rs.bof or rs.eof then
		rs.close
		sql="Select * From ad where "&add&" checked="&true&""
		rs.Open sql,conn, 3,1
		rs.movefirst
		if rs("adwidth")<>0 then adwidth=" width="&rs("adwidth")
		if rs("adheight")<>0 then adheight=" height="&rs("adheight")
		if rs("isflash")=true then
			jsad="<embed src="""&rs("AdPic")&""" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" " & adwidth & adheight &"></embed><br><br>"	
		else
			jsad="<a href=""ShowAd.asp?ads=" & rs("id") & "&url=" & rs("AdUrl") & """ target=""_blank""><img border=0 src=""" & rs("AdPic") & """  alt=""" & rs("AdIntro") & """ " & adwidth & adheight &"></a><br><br>"
		end if
	else
		if rs("adwidth")<>0 then adwidth=" width="&rs("adwidth")
		if rs("adheight")<>0 then adheight=" height="&rs("adheight")
		if rs("isflash")=true then
			jsad="<embed src="""&rs("AdPic")&""" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"" type=""application/x-shockwave-flash"" " & adwidth & adheight &"></embed><br><br>"	
		else
			jsad="<a href=""ShowAd.asp?ads=" & rs("id") & "&url=" & rs("AdUrl") & """ target=""_blank""><img border=0 src=""" & rs("AdPic") & """  alt=""" & rs("AdIntro") & """ " & adwidth & adheight &"></a><br><br>"
		end if
	end if
	rs.close
end if
end function
%>