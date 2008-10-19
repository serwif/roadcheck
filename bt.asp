<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
 <head>
  <title></title>
  <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
  <meta name="ProgId" content="VisualStudio.HTML">
  <meta name="Originator" content="Microsoft Visual Studio .NET 7.1">
  <STYLE>v\:* { BEHAVIOR: url(#default#VML) }
 o\:* { BEHAVIOR: url(#default#VML) }
 .shape { BEHAVIOR: url(#default#VML) }
 </STYLE>
  <script language="javascript">
function Add(){
 var shape=document.createElement("v:shape");
 shape.type="#tooltipshape";
 shape.style.width="150px";
 shape.style.height="150px";
 shape.coordsize="21600,21600";
 shape.fillcolor="infobackground";
 
 var textbox=document.createElement("v:textbox");
 textbox.style.border="1px solid red";
 textbox.style.innerHTML="����";
 shape.appendChild(textbox); 
 
 document.body.appendChild(shape);
}

function VMLPie(pContainer,pWidth,pHeight,pCaption){
 this.Container=pContainer;
 this.Width= pWidth || "400px";
 this.Height=pHeight || "250px";
 this.Caption = pCaption || "VML Chart";
 this.backgroundColor="";
 this.Shadow=false;
 this.BorderWidth=0;
 this.BorderColor=null;
 this.all=new Array();
 this.RandColor=function(){
  
  return "rgb("+ parseInt( Math.random() * 255) +"," +parseInt( Math.random() * 255) +"," +parseInt( Math.random() * 255)+")";
 }
 this.VMLObject=null;
}

VMLPie.prototype.Draw=function(){
 //�����
  var o=document.createElement("v:group");
  o.style.width=this.Width;
  o.style.height=this.Height;
  o.coordsize="21600,21600";
 //���һ��������
  var vRect=document.createElement("v:rect");
  vRect.style.width="34600px"
  vRect.style.height="21600px"
  o.appendChild(vRect);
  
  var vCaption=document.createElement("v:textbox");
  vCaption.style.fontSize="24px";  
  vCaption.style.height="24px"
  vCaption.style.fontWeight="bold";
  vCaption.innerHTML=this.Caption;
  vCaption.style.textAlign="left";
  
  vRect.appendChild(vCaption);
 //���ñ߿��С 
  if(this.BorderWidth){
   vRect.strokeweight=this.BorderWidth;
  }
 //���ñ߿���ɫ
  if(this.BorderColor){
   vRect.strokecolor=this.BorderColor;
  }
 //���ñ�����ɫ
  if(this.backgroundColor){  
   vRect.fillcolor=this.backgroundColor;
  }
 //�����Ƿ������Ӱ
  if(this.Shadow){
   var vShadow=document.createElement("v:shadow");
   vShadow.on="t";
   vShadow.type="single";
   vShadow.color="graytext";
   vShadow.offset="4px,4px";
   vRect.appendChild(vShadow);
  }
  this.VMLObject=o;
  this.Container.appendChild(o);
  
 //��ʼ���ڲ�԰
  var oOval=document.createElement("v:oval");
  oOval.style.width="15000px";
  oOval.style.height="14000px";
  oOval.style.top="4000px";
  oOval.style.left="1000px";
  oOval.fillcolor="#d5dbfb";
  
  //�����ƻ�����3D��Ч���������о�ȷʵ���ÿ��ƣ������ö�����
  //var o3D=document.createElement("o:extrusion");
  var formatStyle=' <v:fill  rotate="t" angle="-135" focus="100%" type="gradient"/>';
  //formatStyle+='<o:extrusion v:ext="view" color="#9cf" on="t" rotationangle="-15"';
  //formatStyle+=' viewpoint="0,34.72222mm" viewpointorigin="0,.5" skewangle="105"';
  //formatStyle+=' lightposition="0,50000" lightposition2="0,-50000"/>';
  formatStyle+='<o:extrusion v:ext="view" backdepth="1in" on="t" viewpoint="0,34.72222mm"   viewpointorigin="0,.5" skewangle="90" lightposition="-50000"   lightposition2="50000" type="perspective"/>';
  oOval.innerHTML=formatStyle;  
  
  o.appendChild(oOval);
  this.CreatePie(o);  
  
}
VMLPie.prototype.CreatePie=function(vGroup){
  var mX=Math.pow(2,16) * 360;
  //��������ǻ�ͼ�εĹؼ� 
  //AE x y width height startangle endangle
  //x y��ʾԲ��λ��
  //width height��״�Ĵ�С
  //startangle endangle�ļ��㷽������
  // 2^16 * ���� 
  
  var vTotal=0;
  var startAngle=0;
  var endAngle=0;
  var pieAngle=0;
  var prePieAngle=0;
  
  var objRow=null;
  var objCell=null;
  
  for(i=0;i<this.all.length;i++){
   vTotal+=this.all[i].Value;
  }
  
  var objLegendRect=document.createElement("v:rect");
  
  var objLegendTable=document.createElement("table");
  objLegendTable.cellPadding=0;
  objLegendTable.cellSpacing=3;
  objLegendTable.width="100%";
  with(objLegendRect){
   style.left="17000px";
   style.top="200px";
   style.width="17350px";
   style.height="21100px";
   fillcolor="#e6e6e6";
   strokeweight="1px";   
  }
  objRow=objLegendTable.insertRow();
  objCell=objRow.insertCell();
  objCell.colSpan="2";
  //objCell.style.border="1px outset";
  objCell.style.backgroundColor="black";
  objCell.style.padding="5px";
  objCell.style.color="window";
  objCell.style.font="caption";
  objCell.innerText="����:"+vTotal;
  
  
  var vTextbox=document.createElement("v:textbox");  
  vTextbox.appendChild(objLegendTable);
  objLegendRect.appendChild(vTextbox);
  
  var vShadow=document.createElement("v:shadow");
  vShadow.on="t";
  vShadow.type="single";
  vShadow.color="graytext";
  vShadow.offset="2px,2px";
  objLegendRect.appendChild(vShadow);
  
  
  vGroup.appendChild(objLegendRect);  

  
  var strAngle="";
  for(i=0;i<this.all.length;i++){ //˳��Ļ���������ͼ
   var vPieEl=document.createElement("v:shape");
   var vPieId=document.uniqueID;
   vPieEl.style.width="15000px";
   vPieEl.style.height="14000px";
   vPieEl.style.top="4000px";
   vPieEl.style.left="1000px";
   vPieEl.coordsize="1500,1400";
   vPieEl.strokecolor="white"; 
   vPieEl.style.fontSize="9pt";
   vPieEl.id=vPieId;
   
   pieAngle= this.all[i].Value / vTotal;
   startAngle+=prePieAngle;
   prePieAngle=pieAngle;
   endAngle=pieAngle; 
   
   //strAngle+=this.all[i].Name +":" +this.all[i].Value+ " Start:"+startAngle +"  End:"+ endAngle +"\n";
   
   vPieEl.path="M 750 700 AE 750 700 750 700 " + parseInt(mX * startAngle) +" " + parseInt(mX * endAngle) +" xe";
  
   vPieEl.title=this.all[i].Name +":"+ parseFloat(endAngle * 100).toFixed(2)+"%" //+this.all[i].TooltipText;
   
   //vPieEl.innerHTML='<v:fill  rotate="t" angle="-135" focus="100%" type="gradient"/>';
   var objFill=document.createElement("v:fill");
   objFill.rotate="t";
   objFill.focus="100%";
   objFill.type="gradient";
   objFill.angle=parseInt( 360 * (startAngle + endAngle /2));
   vPieEl.appendChild(objFill);
   
   var objTextbox=document.createElement("v:textbox");
   objTextbox.border="1px solid black";
   objTextbox.innerHTML=this.all[i].Name +":" + this.all[i].Value ;
   //vPieEl.appendChild(objTextbox);
   
   var vColor=this.RandColor();
   vPieEl.fillcolor=vColor; //������ɫ
   //��ʼ��ͼ��
   objRow=objLegendTable.insertRow();
   objRow.style.height="16px";
   
   var objColor=objRow.insertCell();//��ɫ���
   objColor.style.backgroundColor=vColor;
   objColor.style.width="16px";
   
   objColor.PieElement=vPieId;
   objColor.attachEvent("onmouseover",LegendMouseOverEvent);
   objColor.attachEvent("onmouseout",LegendMouseOutEvent);
   //objColor.onmouseover="LegendMouseOverEvent()";
   //objColor.onmouseout="LegendMouseOutEvent()";
   
   objCell=objRow.insertCell();
   objCell.style.font="icon";
   objCell.style.padding="1px";
   objCell.innerText=this.all[i].Name +":"+this.all[i].Value +"/"+parseFloat(endAngle * 100).toFixed(2)+"%";
   
   vGroup.appendChild(vPieEl);
  }
  
}
VMLPie.prototype.Refresh=function(){

}
VMLPie.prototype.Zoom=function (iValue){
 var vX=21600;
 var vY=21600;
 this.VMLObject.coordsize=parseInt(vX / iValue) +","+parseInt(vY /iValue);
}
VMLPie.prototype.AddData=function(sName,sValue,sTooltipText){

 var oData=new Object();
 oData.Name=sName;
 oData.Value=sValue;
 oData.TooltipText=sTooltipText;
 var iCount=this.all.length;
 this.all[iCount]=oData;

}
VMLPie.prototype.Clear=function(){
 this.all.length=0;
}
function LegendMouseOverEvent(){
 
 var eSrc=window.event.srcElement;
 eSrc.border="1px solid black";
}
function LegendMouseOutEvent(){
 var eSrc=window.event.srcElement;
 eSrc.border="";
}


var objPie=null;
//�����Ǻ�������
function DrawPie(){
 objPie=new VMLPie(document.body,"600px","600px","�˿�ͳ��ͼ");
 //objPie.BorderWidth=3;
 //objPie.BorderColor="blue";
 //objPie.Width="800px";
 //objPie.Height="600px";
 objPie.backgroundColor="#ffffff";
 objPie.Shadow=false;
 
 objPie.AddData("����",50,"");
 objPie.AddData("���",30,"");
 objPie.AddData("��ר",10,"");

 objPie.Draw();
 //alert(document.body.outerHTML);
}

  </script>
 </head>
 <body>
  
 <script language="javascript">
 objPie=new VMLPie(document.body,"330px","330px","ͳ��ͼ");
 objPie.backgroundColor="#ffffff";
 objPie.Shadow=false;

 //objPie.BorderColor="blue";
 objPie.AddData("����",40,"");
 objPie.AddData("��ר",50,"");
 objPie.AddData("��ר",10,"");

 objPie.Draw();

 </script>
 <form action="bing.asp" method="post">
  <p> ��Ŀ1 
    <input name="item1" type="text" id="item1">
  </p>
  <p> ��Ŀ2 
    <input name="item2" type="text" id="item2">
  </p>
  <p> ��Ŀ3 
    <input name="item3" type="text" id="item3">
  </p>
  <p>
    <input type="submit" name="Submit" value="Submit">
  </p>
</form>
 </body>
</html>

