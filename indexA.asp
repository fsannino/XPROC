<%@LANGUAGE="VBSCRIPT"%>
<%
Session.TimeOut=120

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
%>
<script language="javascript">
	//*** ESTA VARIÁVEL RECEBERÁ A CATEGORIA DO USUÁRIO PARA MONTAR O MENU
	var str_CategoriaUsuario = "<%=Session("CatUsu")%>";
</script>
<%
    ls_Script = "<script language=""JavaScript"" src=""Templates/js/indexMenu.js""></script>"
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
<style type="text/css">
<!--
.style6 {
	font-size: 8pt;
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
-->
</style>
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
<script language="JavaScript" type="text/JavaScript">

<!--
    opener.opener = opener;
    opener.close();
//-->
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:mover()" onKeyDown="javascript:ver_tecla()">
<div id="Layer1" style="position:absolute; left:644px; top:217px; width:58px; height:40px; z-index:1; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000"><a href="asp/fale_conosco.asp"><img src="../xproc/imagens/mail.gif" width="73" height="39" border="0"></a></div>
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
        <%'="aaa " & Session("AcessoUsuario")%>
        <font color="#FFFFFF"><%'=Application("Datainicial")%> <b><font size="1" face="Arial">
		<% if Session("CategoriaUsu") = "indexQ.htm" then %>
        	</font></b><span class="style6"><%=Session("Conn_String_Cogest_Gravacao")%></span><b><font size="1" face="Arial">
		<% end if %>
		    </font>
		    <%'=Session("CategoriaUsu")'%>
            </b> </font></p>
    </td>
    <td width="279" valign="top">
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr>
          <td bgcolor="#330099" width="39" valign="middle" align="center">
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img src="../xproc/imagens/voltar.gif" alt=":: Volta" width="30" height="30" border="0"></a>
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center">
            <div align="center"><a href="JavaScript:history.forward()"><img src="../xproc/imagens/avancar.gif" alt=":: Avança" width="30" height="30" border="0"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center">
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000ws10.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Proc')"><img src="../xproc/imagens/favoritos.gif" alt=":: Favorecido" width="30" height="30" border="0"></a></div>
          </td>
        </tr>
        <tr>
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center">
            <div align="center"><a href="javascript:print()"><img src="../xproc/imagens/imprimir.gif" alt=":: Imprime" width="30" height="30" border="0"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center">
            <div align="center"><a href="JavaScript:history.go()"><img src="../xproc/imagens/atualizar.gif" alt=":: Atualiza página" width="30" height="30" border="0"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center">
            <p align="center"><a href="JavaScript:window.close()"><img src="imagens/sair.gif" alt=":: Sair do sistema" width="26" height="24" border="0"></a>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="2" height="20" width="550"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Usu&aacute;rio
      : <%=Session("CdUsuario")%></font></b>
      <%if session("MegaProcesso")<>0 then
    set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & session("MegaProcesso"))
    VALOR=RS("MEPR_TX_DESC_MEGA_PROCESSO")
    %>
    <font face="Verdana" size="2"><b>Mega-Processo Atual : <%=valor%></b></font>
    <%end if%></td>
    <td colspan="2" height="20" width="231"><font face="Verdana" size="2"><b></b></font></td>
  </tr>
</table>
<table border="0" width="77%">
  <tr>
    <td width="8%"></td>
    <td width="97%">&nbsp;
      <p><img src="../xproc/imagens/fundoXProc.jpg" width="692" height="360" border="0">
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

</body>

</html>