<%@ Language=JavaScript %>
<%
var strCountData="1234567890";
var theTime=new Date();

function GetOnline()
{
if (typeof(Application("GuestOnline"))=="undefined")
Application("GuestOnline")="0";
}

function checkGuest()
{
//check Me
if (typeof(Session("test"))=="undefined")
Session("test")="guest";
var strUserName=String(Session("test"));
var strGuestOnline=String(Application("GuestOnline"));
var GuestArray=strGuestOnline.split("\n");
var i;
var iGuestNum;
if (typeof(Session("GuestNum"))=="undefined")
{
for (i=0;i<GuestArray.length;i++)
{
if (GuestArray[i]=="0") break;
}
Session("GuestNum")=i;
}

iGuestNum=Session("GuestNum")*1;
if (iGuestNum>10000)iGuestNum=10000;

var strFormatTime=theTime.getTime();
strFormatTime="0000000000000000000"+strFormatTime;
strFormatTime=strFormatTime.substr(strFormatTime.length-16,16);

GuestArray[iGuestNum]=strFormatTime+strUserName;

strGuestOnline=GuestArray.join("\n");
Application("GuestOnline")=strGuestOnline;

return 1;
}

function GetNumber()
{
//Count Guests on line
var strFormatTime=theTime.getTime()-5*60*1000;
strFormatTime="0000000000000000000"+strFormatTime;
strFormatTime=strFormatTime.substr(strFormatTime.length-16,16);

var strGuestOnline=String(Application("GuestOnline"));
var GuestArray=strGuestOnline.split("\n");

var iGuestCounter=0;
var i;
for (i=0;i<GuestArray.length;i++)
{
if (GuestArray[i].substr(0,16)> strFormatTime)
iGuestCounter++;
else if (GuestArray[i].length>0)
GuestArray[i]="0";
}
strGuestOnline=GuestArray.join("\n");
Application("GuestOnline")=strGuestOnline;
strCountData=iGuestCounter;
return strCountData;
}

function OutPut()
{
var i;
strCountData=""+strCountData;

var strDigits= new Array(
"0","0x3c","0x66","0x66","0x66","0x66","0x66","0x66","0x66","0x66","0x3c", //0
"1","0x30","0x38","0x30","0x30","0x30","0x30","0x30","0x30","0x30","0x30",  //1
"2","0x3c","0x66","0x60","0x60","0x30","0x18","0x0c","0x06","0x06","0x7e",  //2
"3","0x3c","0x66","0x60","0x60","0x38","0x60","0x60","0x60","0x66","0x3c",  //3
"4","0x30","0x30","0x38","0x38","0x34","0x34","0x32","0x7e","0x30","0x78",  //4
"5","0x7e","0x06","0x06","0x06","0x3e","0x60","0x60","0x60","0x66","0x3c",  //5
"6","0x38","0x0c","0x06","0x06","0x3e","0x66","0x66","0x66","0x66","0x3c",  //6
"7","0x7e","0x66","0x60","0x60","0x30","0x30","0x18","0x18","0x0c","0x0c",  //7
"8","0x3c","0x66","0x66","0x66","0x3c","0x66","0x66","0x66","0x66","0x3c",  //8
"9","0x3c","0x66","0x66","0x66","0x66","0x7c","0x60","0x60","0x30","0x1c"); //9

var iCharCount=strCountData.length;
var iCharWidth=8;
var iCharHeight=10;
var theBit;
var theNum;
Response.ContentType ="image/x-xbitmap";
Response.Expires =0;

Response.Write ("#define counter_width "+ iCharWidth*iCharCount+"\r\n");  //м╪пн©М
Response.Write ("#define counter_height "+ iCharHeight+"\r\n");   //м╪пн╦ъ
Response.Write ("static unsigned char counter_bits[]={\r\n");

for (iRow=0;iRow<iCharHeight;iRow++)
for (i=0;i<iCharCount;i++)
{
theBit=strCountData.charAt(i);
for (k=0;k<strDigits.length;k+=(iCharHeight+1))
{
if (strDigits[k]==theBit)break;
}
if (k>=strDigits.length)k=0;
theOffset=k+1;

Response.Write (strDigits[theOffset+iRow]);
Response.Write (",");
}

Response.Write ("};\r\n");

}

GetOnline();
checkGuest();
GetNumber();
OutPut();
%>