<%dim total(7,3)
total(1,0)="中国经营报"
total(2,0)="招聘网"
total(3,0)="51Job"
total(4,0)="新民晚报"
total(5,0)="新闻晚报"
total(6,0)="南方周末"
total(7,0)="羊城晚报"

total(0,1)="#FF0000,1.5,1,2,公司1"'参数1线条的颜色，参数2线条的宽度，参数3线条的类型，参数4转折点的类型,参数5线条名称
total(1,1)=200
total(2,1)=1200
total(3,1)=900
total(4,1)=600
total(5,1)=1222
total(6,1)=413
total(7,1)=800

total(0,2)="#0000FF,1,2,3,公司2"
total(1,2)=400
total(2,2)=500
total(3,2)=1040
total(4,2)=1600
total(5,2)=522
total(6,2)=813
total(7,2)=980

total(0,3)="#004D00,1,1,3,公司3"
total(1,3)=900
total(2,3)=890
total(3,3)=1240
total(4,3)=1300
total(5,3)=722
total(6,3)=833
total(7,3)=1280

%><html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<!--[if !mso]>
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]-->

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<style>
TD {	FONT-SIZE: 9pt}
</style></head>
<body topmargin=5 leftmargin=0 scroll=AUTO>
<%call table2(total,100,90,600,250,3)%>
</body>
</html>

<%
function table2(total,table_x,table_y,all_width,all_height,line_no)
'参数含义(传递的数组，横坐标，纵坐标，图表的宽度，图表的高度,折线条数)
'纯ASP代码生成图表函数2——折线图
'作者：龚鸣(Passwordgm) QQ:25968152 MSN:passwordgm@sina.com Email:passwordgm@sina.com
'本人非常愿意和ASP,VML,FLASH的爱好者在HTTP://topclouds.126.com进行交流和探讨
'版本1.0 最后修改日期 2003-8-11
'非常感谢您使用这个函数，请您使用和转载时保留版权信息，这是对作者工作的最好的尊重。

line_color="#69f"
left_width=70
total_no=ubound(total,1)

temp1=0
for i=1 to total_no
	for j=1 to line_no
		if temp1<total(i,j) then temp1=total(i,j)
	next
next
temp1=int(temp1)
if temp1>9 then
	temp2=mid(cstr(temp1),2,1)
	if temp2>4 then 
		temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+1)*10^(len(cstr(temp1))-1)
	else
		temp3=(int(temp1/(10^(len(cstr(temp1))-1)))+0.5)*10^(len(cstr(temp1))-1)
	end if
else
	if temp1>4 then temp3=10 else temp3=5
end if
temp4=temp3
response.write "<v:rect id='_x0000_s1027' alt='' style='position:absolute;left:"&table_x+left_width&"px;top:"&table_y&"px;width:"&all_width&"px;height:"&all_height&"px;z-index:-1' fillcolor='#9cf' stroked='f'><v:fill rotate='t' angle='-45' focus='100%' type='gradient'/></v:rect>"
for i=0 to all_height-1 step all_height/5
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+length&"px,"&table_y+all_height-length-i&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height-length-i&"px' strokecolor='"&line_color&"'/>"
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+(left_width-15)&"px,"&table_y+i&"px' to='"&table_x+left_width&"px,"&table_y+i&"px'/>"
	response.write ""
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x&"px;top:"&table_y+i&"px;width:"&left_width&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='right'>"&temp4&"</td></tr></table></v:textbox></v:shape>"
	temp4=temp4-temp3/5
next
response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y+all_height&"px' to='"&table_x+all_width+left_width&"px,"&table_y+all_height&"px'/>"
response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width&"px,"&table_y&"px' to='"&table_x+left_width&"px,"&table_y+all_height&"px'/>"

dim line_code
redim line_code(line_no,5)
for i=1 to line_no
	line_temp=split(total(0,i),",")
	line_code(i,1)=line_temp(0)
	line_code(i,2)=line_temp(1)
	line_code(i,3)=line_temp(2)
	line_code(i,4)=line_temp(3)
	line_code(i,5)=line_temp(4)
next
for j=1 to line_no
	for i=1 to total_no-1
			x1=table_x+left_width+all_width*(i-1)/total_no
			y1=table_y+(temp3-total(i,j))*(all_height/temp3)
			x2=table_x+left_width+all_width*i/total_no
			y2=table_y+(temp3-total(i+1,j))*(all_height/temp3)
			response.write "<v:line id=""_x0000_s1025"" alt="""" style='position:absolute;left:0;text-align:left;top:0;z-index:1' from="""&x1&"px,"&y1&"px"" to="""&x2&"px,"&y2&"px"" coordsize=""21600,21600"" strokecolor="""&line_code(j,1)&""" strokeweight="""&line_code(j,2)&""">"
		select case line_code(j,3)
		case 1
		case 2
			response.write "<v:stroke dashstyle='1 1'/>"
		case 3
			response.write "<v:stroke dashstyle='dash'/>"
		case 4
			response.write "<v:stroke dashstyle='dashDot'/>"
		case 5
			response.write "<v:stroke dashstyle='longDash'/>"
		case 6
			response.write "<v:stroke dashstyle='longDashDot'/>"
		case 7
			response.write "<v:stroke dashstyle='longDashDotDot'/>"
		case else
		end select
		response.write "</v:line>"&CHR(13)
		select case line_code(j,4)
		case 1
		case 2
			response.write "<v:rect id=""_x0000_s1027"" style='position:absolute;left:"&x1-2&"px;top:"&y1-2&"px;width:4px;height:4px; z-index:2' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		case 3
			response.write "<v:oval id=""_x0000_s1026"" style='position:absolute;left:"&x1-2&"px;top:"&y1-2&"px;width:4px;height:4px;z-index:1' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		end select
	next
		select case line_code(j,4)
		case 1
		case 2
			response.write "<v:rect id=""_x0000_s1027"" style='position:absolute;left:"&x2-2&"px;top:"&y2-2&"px;width:4px;height:4px; z-index:2' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		case 3
			response.write "<v:oval id=""_x0000_s1026"" style='position:absolute;left:"&x2-2&"px;top:"&y2-2&"px;width:4px;height:4px;z-index:1' fillcolor="""&line_code(j,1)&""" strokecolor="""&line_code(j,1)&"""/>"&CHR(13)
		end select
next
	
for i=1 to total_no
	response.write "<v:line id='_x0000_s1027' alt='' style='position:absolute;left:0;text-align:left;top:0;flip:y;z-index:-1' from='"&table_x+left_width+all_width*(i-1)/total_no&"px,"&table_y+all_height&"px' to='"&table_x+left_width+all_width*(i-1)/total_no&"px,"&table_y+all_height+15&"px'/>"
	response.write ""
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+left_width+all_width*(i-1)/total_no&"px;top:"&table_y+all_height&"px;width:"&all_width/total_no&"px;height:18px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left'>"&total(i,0)&"</td></tr></table></v:textbox></v:shape>"
next

tb_height=30
response.write "<v:rect id='_x0000_s1025' style='position:absolute;left:"&table_x+all_width+20&"px;top:"&table_y&"px;width:100px;height:"&line_no*tb_height+20&"px;z-index:1'/>"
for i=1 to line_no
	response.write "<v:shape id='_x0000_s1025' type='#_x0000_t202' alt='' style='position:absolute;left:"&table_x+all_width+25&"px;top:"&table_y+10+(i-1)*tb_height&"px;width:60px;height:"&tb_height&"px;z-index:1'>"
	response.write "<v:textbox inset='0px,0px,0px,0px'><table cellspacing='3' cellpadding='0' width='100%' height='100%'><tr><td align='left'>"&line_code(i,5)&"</td></tr></table></v:textbox></v:shape>"
	response.write "<v:rect id='_x0000_s1040' alt='' style='position:absolute;left:"&table_x+all_width+80&"px;top:"&table_y+10+(i-1)*tb_height+4&"px;width:30px;height:20px;z-index:1' fillcolor='"&line_code(i,1)&"'><v:fill color2='"&line_code(i,1)&"' rotate='t' focus='100%' type='gradient'/></v:rect>"
next

end function
%>

