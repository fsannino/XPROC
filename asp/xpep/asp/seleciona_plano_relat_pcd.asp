<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
str_Sigla_Plano = "PCD"

on error resume next
	set db_Cogest = Server.CreateObject("ADODB.Connection")
	db_Cogest.Open Session("Conn_String_Cogest_Gravacao")
	
	set db_Cronograma = Server.CreateObject("ADODB.Connection")
	db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

if err.number <. 0 then		
	strMSG = "Ocorreu algum problema com o servidor!"
	Response.redirect "msg_geral.asp?pMsg=" & strMsg & "&pErroServidor=S"
end if	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
<!-- InstanceEndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333; text-decoration: none}
a:hover {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333;  text-decoration: underline}
-->
</style>
<link href="/css/biblioteca.css" rel="stylesheet" type="text/css">
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="Head01" -->

<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaScri" -->
<script language="JavaScript">
	function chamapagina()
	{		
		document.frm1.action="seleciona_plano_relat_pcd.asp?pOrigem=1";
		document.frm1.submit();
	}
		
	function Confirma()
	{			
		document.frm1.action="relat_pcd_x_atividade.asp";          
		document.frm1.submit();		
	}

	function Limpa()
	{
		window.location.href='seleciona_plano_relat_pcd.asp?pOrigem=1'
	}
		
	function mOvr(src,clrOver)
	{
		if (!src.contains(event.fromElement)) 
		{
			src.style.cursor = 'hand';
			src.bgColor = clrOver;
		}
	}
	
	function mOut(src,clrIn) 
	{
		if (!src.contains(event.toElement))
		{
			src.style.cursor = 'default';
			src.bgColor = clrIn;
		}
	}
	
	function mClk(src) 
	{
		if(event.srcElement.tagName=='TD')
		{
			src.children.tags('A')[0].click();
		}
	}
</script>
<!-- InstanceEndEditable -->
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<div id="Layer1" style="position:absolute; left:20px; top:10px; width:134px; height:53px; z-index:1"><img src="../img/000005.gif" alt=":: Logo Sinergia" width="134" height="53" border="0" usemap="#Map2"> 
	  <map name="Map2">
	    <area shape="rect" coords="6,7,129,49">
	  </map>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><table width="780" height="44" border="0" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="583" height="44"><img src="../img/_0.gif" width="1" height="1"></td>
	          <td width="197" height="44"><img src="../../../imagens/000043.gif" width="95" height="44"></td>
	        </tr>
	      </table></td>
	  </tr>
</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td bgcolor="#6699CC">
			<table width="780" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td width="154" height="21"><img src="../img/000002.gif" width="154" height="21"></td>
			    <td width="19" height="21"><img src="../img/000003.gif" width="19" height="21"></td>
			    <td width="202" height="21">
					<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
						<strong>
						</strong>
					</font>
			    </td>
			    <td>&nbsp;</td>
		      </tr>
			</table>
	    </td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="5"><img src="../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="780" height="58" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20" height="39"><img src="../img/_0.gif" width="1" height="1"></td>
        <td width="740" height="39" background="../img/000006.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="11%">&nbsp;</td>
              <td width="13%">&nbsp;</td>
              <td width="61%"><font color="#666666" size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>PLANO DE ENTRADA EM PRODU&Ccedil;&Atilde;O</b></font></td>
              <td width="15%"><a href="../../../indexA_xpep.asp"><img src="../img/botao_home_off_01.gif" alt="Ir para tela inicial" width="34" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td width="20" height="39"><img src="../img/_0.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<!-- InstanceBeginEditable name="corpo" -->
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="10%">&nbsp;</td>
    <td width="77%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo"><strong>Consulta das Atividades - Plano de Conversões de Dados - PCD</strong></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<form method="post" name="frm1" id="frm1">
  <table width="98%" border="0" cellspacing="10">
    <tr> 
      <td width="13%"><div align="right" class="campob">Onda:</div></td>
      <td width="85%"><!--#include file="../includes/inc_Combo_Onda_Rel.asp"-->
      </td>
      <td width="2%">&nbsp;</td>
    </tr>
    <%if Request("selOnda") = "5" OR Request("selOnda") = "7" then%>
    <tr>
      <td class="campob"><div align="right">Fase:</div></td>
      <td><!--#include file="../includes/inc_combo_fases_Rel.asp" --></td>
      <td>&nbsp;</td>
    </tr>
    <%end if%>
	
    <tr> 
      <td class="campob"><div align="right" class="campob">Plano:</div></td>
      <td class="campob"><!--#include file="../includes/inc_Combo_Plano_relat.asp" --></td>
      <td>&nbsp;</td>
    </tr>			
	<tr>
	  <td height="21" valign="top"><div align="right" class="campob">Atividades:</div></td>
	  <td><!--#include file="../includes/inc_combo_tarefas_nivel1_Rel.asp" --></td>
	  <td>&nbsp;</td>
	</tr>	
  </table>
  <table width="625" border="0" align="center">
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><input name="hidTpRel" type="hidden" id="hidTpRel" value="<%=str_TpRel%>"></td>
    </tr>
    <tr>
      <td width="26"><a href="javascript:Confirma()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
      <td width="26"><b></b></td>
      <td width="195"><a href="javascript:Limpa()"><img src="../img/limpar_01.gif" width="85" height="19" border="0"></a>        <!--<a href="javascript:Limpa();"><img src="../img/limpar_01.gif" width="85" height="19" border="0"></a>--></td>
      <td width="27"></td>
      <td width="50">&nbsp;</td>
      <td width="28"></td>
      <td width="26">&nbsp;</td>
      <td width="159"></td>
    </tr>
  </table>
</form>
<%
  db_Cronograma.Close
  db_Cogest.close
  set db_Cronograma = Nothing
  set db_Cogest = Nothing
  %>
<!-- InstanceEndEditable -->
    <table width="200" border="0" align="center">
<tr>	
	<td height="10" width="780"></td>
</tr>
<tr>
	<td width="780">			
		<p width="780" align="center"><img src="../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>
