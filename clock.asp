<html>
<head>
<meta  http-equiv="Content-Type"  content="text/html;  charset=gb2312">
<meta  name="GENERATOR"  content="Microsoft  FrontPage  4.0">
<meta  name="ProgId"  content="FrontPage.Editor.Document">
<title>make标准服务器时钟</title>
<script  language="javascript">
<!--  
var  timerID  =  null
var  timerRunning  =  false  
function  stopclock(){
if(timerRunning)
clearTimeout(timerID)
timerRunning  =  false
}  
function  startclock(){
stopclock()
showtime()
}
<%
response.write "var  years  =  "&year(now)
response.write chr(13)
response.write "var  months  =  "&month(now)
response.write chr(13)
response.write "var  days  =  "&day(now)
response.write chr(13)
response.write "var  hours  =  "&hour(now)
response.write chr(13)
response.write "var  minutes  =  "&minute(now)
response.write chr(13)
response.write "var  seconds  =  "&second(now)
response.write chr(13)
response.write chr(13)
response.write "function  showtime(){"
response.write chr(13)
response.write "if (seconds != -1){"
response.write chr(13)
response.write "seconds ++"
response.write chr(13)
response.write "}"
response.write chr(13)
response.write "if (seconds == 60) {"
response.write chr(13)
response.write "seconds = 0"
response.write chr(13)
response.write "minutes++"
response.write chr(13)
response.write "}"
response.write chr(13)
response.write "if (minutes == 60) {"
response.write chr(13)
response.write "minutes = 0"
response.write chr(13)
response.write "hours++"
response.write chr(13)
response.write "}"
response.write chr(13)
response.write "if (hours == 24){"
response.write chr(13)
response.write "hours = 0"
response.write chr(13)
response.write "}"
'response.write "var  timeValue  =  """&"""  +  ((hours  >  12)  ?  hours  -  12  :  hours)"
response.write "var  timeValue  =  """&"""  +  ((hours  >  12)  ?  hours  :  hours)"
response.write chr(13)
response.write "timeValue  +=  ((minutes  <  10)  ?  """&":0"&"""  :  "&""":"""&")  +  minutes"
response.write chr(13)
response.write "timeValue  +=  ((seconds  <  10)  ?  """&":0"&"""  :  "&""":"""&")  +  seconds"
response.write chr(13)
'response.write "timeValue  +=  (hours  >=  12)  ?  "  &"""P.M."""&  "  :  "  &""" A.M."""
'response.write "timeValue =timeValue + years + """+"年"&""" + months + days "
response.write "timeValue = years + ((months  <  10)  ?  """&"-0"&"""  :  "&"""-"""&") + months + ((days  <  10)  ?  """&"-0"&"""  :  "&"""-"""&") + days + """+" "&""" + timeValue"
response.write chr(13)
response.write "document.clock.face.value  =	timeValue"
response.write chr(13)
'//timeValue  
response.write "timerID  =  setTimeout("&"""showtime()"""&",1000)"
response.write chr(13)
response.write "timerRunning  =  true"
response.write chr(13)
response.write "}"
response.write chr(13)
response.write "//-->"
%>
</script>  
</head>
<body  bgcolor="#3366cc"  onload="startclock()">
<form  name="clock"  onsubmit="0">
<input  type="text"  name="face"  size="19">
</form>  
</body>
</html>