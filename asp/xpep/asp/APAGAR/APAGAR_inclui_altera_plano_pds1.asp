<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_RespLegado = ""
str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
str_RespLegado = str_RespLegado & " USMA_CD_USUARIO "
str_RespLegado = str_RespLegado & " , USMA_TX_NOME_USUARIO "
str_RespLegado = str_RespLegado & " FROM dbo.USUARIO_MAPEAMENTO "
str_RespLegado = str_RespLegado & " Where USMA_TX_MATRICULA <> 0"
str_RespLegado = str_RespLegado & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_RespLegado = db_Cogest.Execute(str_RespLegado)
%>
<html>
<!-- InstanceBegin template="/Templates/BASICO_XPEP_01.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
	<!-- InstanceBeginEditable name="doctitle" -->
	<title>SINERGIA # XPROC # Processos de Negócio</title>
	<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
	<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
	<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaSci" -->
	<script src="../js/troca_lista.js" language="javascript"></script>
	<script src="../js/global.js" language="javascript"></script>	
<!-- InstanceEndEditable -->
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../Funcao/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../../indexA.asp"><img src="../../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
      
    <td colspan="3" height="20"><!-- InstanceBeginEditable name="Botao_01" -->
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../../../imagens/continua_F02.gif" width="24" height="24" border="0"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
      <!-- InstanceEndEditable --></td>
  </tr>
</table>     
<!-- InstanceBeginEditable name="Corpo_Princ" -->
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="6%">&nbsp;</td>
    <td width="81%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo"><table width="94%" border="0" cellpadding="0" cellspacing="7">
        <tr> 
          <td width="29%"><div align="right" class="subtitulob">Onda:</div></td>
          <td width="71%" class="subtitulo">SEA</td>
        </tr>
        <tr> 
          <td><div align="right"><span class="subtitulob">Plano:</span></div></td>
          <td class="subtitulo">Plano de Desligamento de Sistemas Legados - PDS </td>
        </tr>
      </table></td>
    <td><table width="75%" border="0">
        <tr>
          <td class="campo"><div align="center">A&ccedil;&atilde;o</div></td>
        </tr>
        <tr>
          <td bgcolor="#EEEEEE"> 
            <div align="center" class="campob">Incluir</div></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="75%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr> 
    <td width="17%" bgcolor="#EEEEEE"> <div align="right" class="campo">Linha:</div></td>
    <td colspan="3"><%="A"%></td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Respons&aacute;vel:</div></td>
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Data In&iacute;cio:</div></td>
    <td width="21%"><%="A"%></td>
    <td width="20%" bgcolor="#EEEEEE"> <div align="right" class="campo">Data de 
        T&eacute;rmino:</div></td>
    <td width="33%"><%="A"%></td>
  </tr>
</table>
<form name="frm1" method="post" action="">
  <table width="98%" border="0">
  
  <td class="campo">&nbsp;</td>
      <td class="campob"><div align="right">Codigo do Sistema:</div></td>
      <td><input type="text" name="txtCDSistema"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
  </tr>
  
  <td class="campo">&nbsp;</td>
      <td class="campob"><div align="right">Nome do Sistema:</div></td>
      <td><input type="text" name="txtNomeSistema"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
  </tr>
  
  <tr>
    <td class="campo">&nbsp;</td>
    <td height="25" class="campob"><div align="right">A&ccedil;&otilde;es Corretivas/Conting&ecirc;ncias:</div></td>
    <td class="campo"><!--#include file="../includes/inc_Combo_Desligamento.asp" --></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
     <td colspan="5"><!--#include file="../includes/inc_lista_Responsavel_Legado.asp" --></td>
  </tr>
  
  <table width="98%" border="0">
  
  <td class="campo">&nbsp;</td>
      <td class="campob"><div align="right">Gerência Téc Resp pelo Legado:</div></td>
      <td><input type="text" name="txtGerTecRespLegado"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
  </tr>  
  
  </table>
  </table>
</form>
<!-- InstanceEndEditable -->
</body>

<!-- InstanceEnd --></html>
