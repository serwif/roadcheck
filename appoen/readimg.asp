<%
Class qswhImg
dim aso
Private Sub Class_Initialize
set aso=CreateObject("Adodb.Stream")
aso.Mode=3 
aso.Type=1 
aso.Open 
End Sub

Private Sub Class_Terminate
set aso=nothing
End Sub

Private Function Bin2Str(Bin)
Dim I, Str
For I=1 to LenB(Bin)
clow=MidB(Bin,I,1)
if ASCB(clow)<128 then
Str = Str & Chr(ASCB(clow))
else
I=I+1
if I <= LenB(Bin) then Str = Str & Chr(ASCW(MidB(Bin,I,1)&clow))
end if
Next 
Bin2Str = Str
End Function

Private Function Num2Str(num,base,lens)
'qiushuiwuhen (2002-8-12)
dim ret
ret = ""
while(num>=base)
ret = (num mod base) & ret
num = (num - num mod base)/base
wend
Num2Str = right(string(lens,"0") & num & ret,lens)
End Function

Private Function Str2Num(str,base)
'qiushuiwuhen (2002-8-12)
dim ret
ret = 0
for i=1 to len(str)
ret = ret *base + cint(mid(str,i,1))
next
Str2Num=ret
End Function

Private Function BinVal(bin)
'qiushuiwuhen (2002-8-12)
dim ret,i
ret = 0
for i = lenb(bin) to 1 step -1
ret = ret *256 + ascb(midb(bin,i,1))
next
BinVal=ret
End Function

Private Function BinVal2(bin)
'qiushuiwuhen (2002-8-12)
dim ret,i
ret = 0
for i = 1 to lenb(bin)
ret = ret *256 + ascb(midb(bin,i,1))
next
BinVal2=ret
End Function

Function getImageSize(filespec) 
'qiushuiwuhen (2002-9-3)
dim ret(3),bFlag
aso.LoadFromFile(filespec)
bFlag=aso.read(3)
select case hex(binVal(bFlag))
case "4E5089":
aso.read(15)
ret(0)="PNG"
ret(1)=BinVal2(aso.read(2))
aso.read(2)
ret(2)=BinVal2(aso.read(2))
case "464947":
aso.read(3)
ret(0)="GIF"
ret(1)=BinVal(aso.read(2))
ret(2)=BinVal(aso.read(2))
case "535746":
aso.read(5)
binData=aso.Read(1)
sConv=Num2Str(ascb(binData),2 ,8)
nBits=Str2Num(left(sConv,5),2)
sConv=mid(sConv,6)
while(len(sConv)<nBits*4)
binData=aso.Read(1)
sConv=sConv&Num2Str(ascb(binData),2 ,8)
wend
ret(0)="SWF"
ret(1)=int(abs(Str2Num(mid(sConv,1*nBits+1,nBits),2)-Str2Num(mid(sConv,0*nBits+1,nBits),2))/20)
ret(2)=int(abs(Str2Num(mid(sConv,3*nBits+1,nBits),2)-Str2Num(mid(sConv,2*nBits+1,nBits),2))/20)
dim p1
case "FFD8FF":
do 
do: p1=binVal(aso.Read(1)): loop while p1=255 and not aso.EOS
if p1>191 and p1<196 then exit do else aso.read(binval2(aso.Read(2))-2)
do:p1=binVal(aso.Read(1)):loop while p1<255 and not aso.EOS
loop while true
aso.Read(3)
ret(0)="JPG"
ret(2)=binval2(aso.Read(2))
ret(1)=binval2(aso.Read(2))
case else:
if left(Bin2Str(bFlag),2)="BM" then
aso.Read(15)
ret(0)="BMP"
ret(1)=binval(aso.Read(4))
ret(2)=binval(aso.Read(4))
else
ret(0)=""
end if
end select
ret(3)="width=""" & ret(1) &""" height=""" & ret(2) &""""
getimagesize=ret
End Function
End Class
%>