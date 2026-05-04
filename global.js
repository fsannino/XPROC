//*** função utilizada pelo FAQ
function MM_changePropOO(objName,x,theProp,theValue)
{ 			
  var obj = MM_findObj(objName);
  var obj2 = MM_findObj(theValue);
  if (obj && obj2 && (theProp.indexOf("style.")==-1 || obj.style &&  obj2.style )) eval("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
}		

		
function MM_swapImgRestore() 
{ 
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}


function MM_preloadImages() 
{ 
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}


function MM_findObj(n, d) 
{ 
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
  d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}


function MM_swapImage() 
{
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}


//*** UTILIZADO PELO MENU		
function MM_showHideLayers() 
{
   var i,p,v,obj,args=MM_showHideLayers.arguments;
   for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
   if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
   obj.visibility=v; }
}


function Limpa()
{
	var strOpcao = document.forms[0].txtOpcao.value;
				
	//*** ALTERAÇÃO ***
	if (strOpcao == 'A')
	{
		var strCodPergResp = document.forms[0].strCodPergResp.value;
		document.forms[0].action = 'inc_Perg_resp_manut.asp?vCdPerResp=' + strCodPergResp;
		document.forms[0].submit();
	}
	else //*** INCLUSÃO ***
	{
		document.forms[0].reset();
		strPalChavTot = '';
	}
}


//*** UTILIZADO PELO MENU
function MM_reloadPage(init) 
{  
   //reloads the window if Nav4 resized
   if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) 
	{
   	document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
   else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}

MM_reloadPage(true);