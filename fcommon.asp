<%
'==============================================================================
'FCOMMON.ASP
'����������
'Copyright by WuQiang/Monday Studio since 2000, all rights reserved world wide.
'==============================================================================

dim WINNT_CHINESE
WINNT_CHINESE=(len("����һ")=3)

dim MailServerUserName
dim MailServerPassword
dim MailServer

dim sysconstr
sysconstr="DBQ="&server.mappath("db2006.mdb")&";DRIVER={Microsoft Access Driver (*.mdb)};"


dim UploadPath'hjhedit20050909���廨������ͼƬ�Ĵ��λ��
UploadPath =server.mappath("\photo\")

dim ajlb_len,ajlb_x_len,ajlb_len0,ajlb_len1,ajlb_len2,ajlb_len3,ajlb_len4,ajlb_len5,ajlb_str1,ajlb_str2,ajlb_str3,ajlb_str4,ajlb_str5'�������೤�ȶ���
ajlb_len0=10'���濨�����ܳ���,�ݶ������濨������4��
ajlb_len1=2'���濨��ǰ2λ,��8λΪ0
ajlb_len2=4'1���ǰ4λ,��6λΪ0,����ǰ2λ�����濨
ajlb_len3=6'2���ǰ6λ,��4λΪ0,����ǰ4λ����1��
ajlb_len4=8'3���ǰ8λ,��2λΪ0,����ǰ6λ����2��
ajlb_len5=10'4���ǰ10λ,����ǰ8λ����3��
ajlb_str1="00000000"'���濨�Ľ�β
ajlb_str2="000000"'1��Ľ�β
ajlb_str3="0000"'2��Ľ�β
ajlb_str4="00"'3��Ľ�β
ajlb_str5=""'4��Ľ�β
dim fxlb_len0,fxlb_len1'������𳤶ȶ���
fxlb_len0=8'��������ܳ���
fxlb_len1=8'����������,ǰ6λ������С��
dim tjlb_len,tjlb_x_len,tjlb_len0,tjlb_len1,tjlb_len2,tjlb_len3,tjlb_len4,tjlb_len5,tjlb_str1,tjlb_str2,tjlb_str3,tjlb_str4,tjlb_str5'ͳ�Ʊ�����Ŀ��𳤶ȶ���
tjlb_len0=10'ͳ�Ʊ�������ܳ���,�ݶ�������������4��
tjlb_len1=2'ͳ�Ʊ����ǰ2λ,��8λΪ0
tjlb_len2=4'ͳ�Ʊ����1���ǰ4λ,��6λΪ0,����ǰ2λ������
tjlb_len3=6'ͳ�Ʊ����2���ǰ6λ,��4λΪ0,����ǰ4λ����1��
tjlb_len4=8'ͳ�Ʊ����3���ǰ8λ,��2λΪ0,����ǰ6λ����2��
tjlb_len5=10'ͳ�Ʊ����4���ǰ10λ,����ǰ8λ����3��
tjlb_str1="00000000"'ͳ�Ʊ���Ľ�β
tjlb_str2="000000"'ͳ�Ʊ���1��Ľ�β
tjlb_str3="0000"'ͳ�Ʊ���2��Ľ�β
tjlb_str4="00"'ͳ�Ʊ���3��Ľ�β
tjlb_str5=""'ͳ�Ʊ���4��Ľ�β
tjlb_len=2
tjlb_x_len=4
dim unit_len0,unit_len1,unit_len2,unit_len3,unit_str1,unit_str2,unit_str3'��λ���೤�ȶ���
unit_len0=8'��λ�ܳ���
unit_len1=4'��λ(ʡ���͸�����)��ǰ4λ,��4λΪ0
unit_len2=6'��λ(���о�,�־�)��ǰ6λ,��2λΪ0,����ǰ4λ���������ʡ��
unit_len3=8'��λ(�ɳ���)��ǰ8λ,����ǰ6λ��������
unit_str1="0000"'��λ����Ľ�β
unit_str2="00"'��λ����Ľ�β
unit_str3=""'��λС��Ľ�β

sub noRightClick()
%>
<script language=JavaScript>
<!--
var message="";
///////////////////////////////////
function clickIE() {if (document.all) {(message);return false;}}
function clickNS(e) {if 
(document.layers||(document.getElementById&&!document.all)) {
if (e.which==2||e.which==3) {(message);return false;}}}
if (document.layers) 
{document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;}
else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;}
document.oncontextmenu=new Function("return false")
// --> 
</script>
<%
end sub

function skincolor()
  select case session("skin")
    case "orange"
      skincolor="#ffa500"
    case "green"
      skincolor="#8DCC1E"
    case else
      skincolor="#569BE8"
  end select
end function


function GetH(MM)'���Сʱ��
  dim Hour,Min 
  GetH=int(MM/60)
end function

function GetM(MM)'��÷�
  GetM=MM mod 60
end function

function lightskincolor()
  select case session("skin")
    case "green"
      lightskincolor="#e3ff85"
    case "orange"
      lightskincolor="#e6c06c"
    case else
      lightskincolor="#7bc8ff"
  end select
end function

function datetostr(sday)
  dim s,Jyear,Jmonth,Jday
  s=formatdatetime(sday,2)
  'datetostr=left(s,4)+mid(s,6,2)+right(s,2)
  Jyear=cstr(year(sday))
  if month(s)<10 then 
	Jmonth="0"+cstr(month(s))
  else
	Jmonth=cstr(month(s))
  end if
  if day(s)<10 then
	Jday="0"+cstr(day(s))
  else
	Jday=cstr(day(s))
  end if
  s=Jyear+Jmonth+Jday
  datetostr=s
end function

function todatestr(sday)
  if len(trim(sday))=0 then 
    todatestr=""
  else
    todatestr=left(sday,4)+"-"+mid(sday,5,2)+"-"+right(sday,2)
  end if
end function

function timetostr(stime)
  dim s
  s=FormatTime(stime)
  timetostr=left(s,2)+mid(s,4,2)
end function

function totimestr(stime)
  totimestr=left(stime,2)+":"+mid(stime,3,2)
end function

function nowstr()
  dim s
  s=formatdatetime(date,2)
  'nowstr=left(s,4)+right(s,2)
  nowstr=datetostr(s)
end function

'ȡ�ַ������ȣ�һ�������������ַ�
function strLength(str)
  if WINNT_CHINESE then
    dim l, t, c, i

    l=len(str)
    t=l
    for i=1 to l
      c=asc(mid(str,i,1))
      if c<0 then
		c=c+65536
	  end if
      if c>255 then
        t=t+1
      end if
    next
    strLength=t
  else
    strLength=len(str)
  end if
end function

'�ж�һ��������EMAIL��ַ��ʽ���Ƿ���Ч
function isEmailValid(em)
  Dim goby
  goby = True 'Initializing goby to False
  'if the len is less than 5 then it can't be an email
  '(i.e.: a@a.c)
  If Len(em) <= 5 Then
  goby = False
  End If
  If InStr(1, em, "@", 1) < 2 Then
  'If we find one and only one @, then the
  'email address is good to go.
  goby = False
  Else
  If InStr(1,em, ".", 1) < 4 Then
    'Must have a '.' too
    goby = False
  End If
  End If
  isEmailValid = goby
end function

'�ӳ�
Sub delaySecond(DelaySeconds)
  dim SecCount, Sec1, Sec2
  SecCount = 0
  Sec2 = 0
  While SecCount<DelaySeconds + 1
  Sec1 = Second(Time())
  If Sec1 <> Sec2 Then
    Sec2 = Second(Time())
     SecCount = SecCount + 1
  End If
  Wend
End Sub

'���sql�ַ������Ƿ��е����ţ��������ת��
function CheckStr(str)
  dim tstr,l,i,ch

  l=len(str)
  for i=1 to l
  ch=mid(str,i,1)
  if ch="'" then tstr=tstr+"'"
  tstr=tstr+ch
  next
  CheckStr=tstr
end function

'���ַ����е������ַ�ת��ΪHTML�﷨
function htmlencode(str,w)
  dim result, l, i, j
  if isnull(str) then
    htmlencode=""
    exit function
  end if
  l=len(str)
  result=""
  j=0
  for i = 1 to l
    select case mid(str,i,1)
      case "<"
        result=result+"&lt;"
      case ">"
        result=result+"&gt;"
      case chr(34)
        result=result+"&quot;"
      case "&"
        result=result+"&amp;"
      case chr(13)
        result=result+"<br>"
      case chr(9)
        result=result+"&nbsp;&nbsp;&nbsp;&nbsp;"
      case chr(32)
        'result=result+"&nbsp;"
        if i+1<=l and i-1>0 then
          if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then
            result=result+"&nbsp;"
          else
            result=result+" "
          end if
        else
          result=result+"&nbsp;"
        end if
      case else
        result=result+mid(str,i,1)
    end select
    if w <> 0 then
      if mid(str,i,1)=chr(13) then
        j = 0
      else
        if Len(Hex(asc(mid(str,i,1))))>2 then
          j = j + 2
        else
          j = j + 1
        end if
        if j = w or j+1=w then
          result = result + "<br>"
          j = 0
        end if
      end if
    end if
  next
  htmlencode=result
end function

'ͳ����������
function CurrentOnlineUsersCount()
  CurrentOnlineUsersCount = rstonlineusers.RecordCount
end function

'��ʽ��ʱ��
function FormatTime(t)
  FormatTime=formatdatetime(t,4)+right(formatdatetime(t,3),3)
end function

'�Ѻõ�������ʾ��ʽ
function GetFriendlyDateFormat(d)
  dim t

  if datediff("d",date,d) = 0 then
  t = "����" & formattime(d)
  elseif datediff("d",date,d) = -1 then
  t = "����" & formattime(d)
  elseif datediff("d",date,d) = -2 then
  t = "ǰ��" & formattime(d)
  elseif datediff("d",date,d) = -3 then
  t = "��ǰ��" & formattime(d)
  else
  t = formatdatetime(d,2)+" "+formattime(d)
  end if
  GetFriendlyDateFormat = t
end function

'e version
function GetFriendlyDateEnglishFormat(d)
  dim t

  if datediff("d",date,d) = 0 then
  t = "Today " & formattime(d)
  elseif datediff("d",date,d) = -1 then
  t = "Yestoday " & formattime(d)
  'elseif datediff("d",date,d) = -2 then
  't = "ǰ��" & formattime(d)
  'elseif datediff("d",date,d) = -3 then
  't = "��ǰ��" & formattime(d)
  else
  t = formatdatetime(d,2)+" "+formattime(d)
  end if
  GetFriendlyDateEnglishFormat = t
end function

'����ַ���
Function RndString ()
    Dim s, i
    Randomize
    s = ""
    For i = 1 To 6
      if rnd < 0.8 then
        s = s + chr(int(rnd*26)+65)
      else
        s = s + chr(int(rnd*10)+48)
      end if
    Next
    RndString = s
End Function

Function yearday (y)
    If y Mod 4 = 0 And y Mod 100 = 0 Then
        If y Mod 400 = 0 Then
            yearday = 366
        Else
            yearday = 365
        End If
    ElseIf y Mod 4 = 0 Then
        yearday = 366
    Else
        yearday = 365
    End If
End Function

Function monthday (y, m)
    Select Case m
        Case 1, 3, 5, 7, 8, 10, 12
            monthday = 31
        Case 4, 6, 9, 11
            monthday = 30
        Case Else
            If yearday(y) = 366 Then
                monthday = 29
            Else
                monthday = 28
            End If
    End Select
End Function

function getFileExtName(fileName)'ȡ�ú�׺��
    dim pos
    pos=instrrev(filename,".")
    if pos> 0 then 
      getFileExtName=mid(fileName,pos+1)
    else
      getFileExtName=""
    end if
end function 

function getFileName(fileName)'ȡ�ú�׺��
    dim pos
    pos=instrrev(filename,"\")
    if pos> 0 then 
      getFileName=mid(fileName,pos+1)
    else
      getFileName=""
    end if
end function 

%>