<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--

function Confirma2() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega-Processo é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href='relacao_ativ_emp.asp?selAtiv='+document.frm1.selAtividade.value
	 }
 }

function carrega_txt(fbox) {
document.frm1.txtEmpSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtEmpSelecionada.value = document.frm1.txtEmpSelecionada.value + "," + fbox.options[i].value;
   }
}

function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}
function Confirma() 
{ 
     {
 	  carrega_txt(document.frm1.list2);
	  document.frm1.submit();
	 }
}

function Limpa(){
	document.frm1.reset();
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<script language="javascript" src="Planilhas/js/troca_lista.js"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../imagens/novo_registro_02.gif','../imagens/atualiza_02.gif','../imagens/continua_F02.gif','../imagens/continua2_F02.gif')">
<form name="frm1" method="post" action="gera_rel_megaemp.asp">
              <input type="hidden" name="txtEmpSelecionada">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Montar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="90%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="6%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
      <td width="68%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
      <td width="68%"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Decomposição
        por Mega-Processo / Empresa</font></td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
      <td width="68%"> 
        <input type="hidden" name="txtOpc" value="<%=str_Opc%>">
      </td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Mega-Processo</b></font></td>
      <td width="68%"> 
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>xxx</b></font>
      </td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
      <td width="68%"><input type="text" name="selMegaProcesso" size="10"></td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Selecione
        a Empresa / Unidade para Visualização do Relatório</b></font></td>
      <td width="68%">&nbsp;<select size="1" name="txtEmpSelecionada">
          <option value="0">== Selecione a Empresa / Unidade ==</option>
        </select> </td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%" valign="top"></td>
      <td width="68%"> 
      </td>
    </tr>
    
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--

function Confirma2() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega-Processo é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href='relacao_ativ_emp.asp?selAtiv='+document.frm1.selAtividade.value
	 }
 }

function carrega_txt(fbox) {
document.frm1.txtEmpSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtEmpSelecionada.value = document.frm1.txtEmpSelecionada.value + "," + fbox.options[i].value;
   }
}

function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}
function Confirma() 
{ 
     {
 	  carrega_txt(document.frm1.list2);
	  document.frm1.submit();
	 }
}

function Limpa(){
	document.frm1.reset();
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<script language="javascript" src="Planilhas/js/troca_lista.js"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../imagens/novo_registro_02.gif','../imagens/atualiza_02.gif','../imagens/continua_F02.gif','../imagens/continua2_F02.gif')">
<form name="frm1" method="post" action="gera_rel_megaemp.asp">
              <input type="hidden" name="txtEmpSelecionada">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Montar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="90%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="6%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
      <td width="68%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
      <td width="68%"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Decomposição
        por Mega-Processo / Empresa</font></td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
      <td width="68%"> 
        <input type="hidden" name="txtOpc" value="<%=str_Opc%>">
      </td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Mega-Processo</b></font></td>
      <td width="68%"> 
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>xxx</b></font>
      </td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
      <td width="68%"><input type="text" name="selMegaProcesso" size="10"></td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Selecione
        a Empresa / Unidade para Visualização do Relatório</b></font></td>
      <td width="68%">&nbsp;<select size="1" name="txtEmpSelecionada">
          <option value="0">== Selecione a Empresa / Unidade ==</option>
        </select> </td>
    </tr>
    <tr> 
      <td width="6%">&nbsp;</td>
      <td width="26%" valign="top"></td>
      <td width="68%"> 
      </td>
    </tr>
    
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
