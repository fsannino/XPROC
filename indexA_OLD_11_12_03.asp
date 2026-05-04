<%@LANGUAGE="VBSCRIPT"%> 
 
<%
Session.TimeOut=120

'response.write "AAAA"
'response.write Session("Conn_String_Cogest_Gravacao")
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

   'response.write " teste "
   'response.write Session("CatUsu")
   Select Case Session("CatUsu")
   Case "indexA.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexA.js""> </script>"
   Case "indexB.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexB.js""> </script>"
   Case "indexC.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexC.js""> </script>"
   Case "indexD.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexD.js""> </script>"
   Case "indexE.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexE.js""> </script>"
   Case "indexF.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexF.js""> </script>"	  
   Case "indexG.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexG.js""> </script>"
   Case "indexH.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexH.js""> </script>"
   Case "indexP.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexP.js""> </script>"
   Case "indexQ.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexQ.js""> </script>"
   Case "indexV.js"
      ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexV.js""> </script>"
   end Select
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Sistema de Cadastro</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK REL="SHORTCUT ICON" href="http://JOAO/XPROC/imagens/Wrench.ico">
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
</head>
<script>
function Fecha()
{
alert('Você está saindo do ambiente do Aplicativo...Obrigado por utilizar o X-PROC');
}
function mover()
{
window.moveTo(0,0);
}
function ver_tecla()
{
var a = event.keyCode;
if(a==16){
alert('Propriedade SINERGIA @ 2003');
alert(event.width);

}
}
</script>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:mover()" onKeyDown="javascript:ver_tecla()">
<div id="Layer1" style="position:absolute; left:644px; top:217px; width:58px; height:40px; z-index:1; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000"><a href="asp/fale_conosco.asp"><img src="imagens/mail.gif" width="73" height="39" border="0"></a></div>
<%=ls_Script%>
<script type= "text/javascript" language= "JavaScript">
<!--
goMenus();
//-->
</script>

<table width="783" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="136" height="20"><%'=session("MegaProcesso")%><font color="#FFFFFF">&nbsp;</font><%'=Application("Totalvisitas")%></td>
    <td width="362" height="60" valign="middle" colspan="2"> <p align="left"> 
        <%'=Session("AcessoUsuario")%>
        <font color="#FFFFFF"><%=Application("Datainicial")%> <b><font size="1" face="Arial">
        <%=Session("Conn_String_Cogest_Gravacao")%></font></b><b><%'=Session("CategoriaUsu")%>
        </b> </font></p>
    </td>
    <td width="279" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="imagens/voltar.gif"></a>
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <p align="center">&nbsp;
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="2" height="20" width="550"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Usu&aacute;rio 
      : <%=Session("CdUsuario")%></font></b>  / 
      <%if session("MegaProcesso")<>0 then
    set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & session("MegaProcesso"))
    VALOR=RS("MEPR_TX_DESC_MEGA_PROCESSO")
    %>
    <font face="Verdana" size="2"><b>Mega-Processo Atual : <%=valor%></b></font>
    <%end if%>
      <a href="asp/Teste_Paginacao.asp?whichpage=1&pagesize=10">exemplo_pagina&ccedil;&atilde;o</a></td>
    <td colspan="2" height="20" width="231"><font face="Verdana" size="2"><b></b></font></td>
  </tr>
</table>
<table border="0" width="77%">
  <tr> 
    <td width="8%"></td>
    <td width="97%">&nbsp; 
      <p><img border="0" src="imagens/fundoXProc.jpg"> 
    </td>
    <td width="97%" valign="top">&nbsp;</td>
  </tr>
  <tr> 
    <td width="8%">
      <p style="margin-top: 0; margin-bottom: 0"></td>
    <td width="97%"> 
      <table width="600" border="0" cellspacing="0" cellpadding="0" align="right">
        <tr> 
          <td width="15">
            <p style="margin-top: 0; margin-bottom: 0"></td>
          <td width="98">
            <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
          </td>
          <td width="145"> 
            <div align="right">
              <p style="margin-top: 0; margin-bottom: 0"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Desenvolvido 
              por :</font></div>
          </td>
          <td width="342"> 
            <div align="left">
              <p style="margin-top: 0; margin-bottom: 0"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Gest&atilde;o 
              do Conhecimento</b></font></div>
          </td>
        </tr>
      </table>
    </td>
    <td width="97%">
      <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
    </td>
  </tr>
</table>
</html>