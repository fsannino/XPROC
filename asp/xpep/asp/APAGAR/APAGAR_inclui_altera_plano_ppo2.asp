<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

set db_Cronograma = Server.CreateObject("ADODB.Connection")
db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

str_Acao = Request("pAcao")
int_Cd_ProjetoProject = Request("pCdProjProject")
int_Id_TarefaProject = Request("pTArefa")
int_CD_Onda = request("pOnda")
int_ResData = request("pResData")
int_Plano = request("pPlano")

'Response.write "Teste " & int_Plano
'Response.END

if str_Acao = "I" then
   str_Acao = "Inclusão"   
else
   str_Acao = "Alteração"   
end if

'response.Write(str_Acao)  & "<p>"
'response.Write(int_Cd_ProjetoProject)  & "<p>"
'response.Write(int_Id_TarefaProject)  & "<p>"

'==== ENCONTRA DESCRIÇÃO DA ONDA ==================================================
str_Sql_Onda = ""
str_Sql_Onda = str_Sql_Onda & " Select ONDA_TX_DESC_ONDA "
str_Sql_Onda = str_Sql_Onda & " from ONDA "
str_Sql_Onda = str_Sql_Onda & " where ONDA_CD_ONDA = " & int_CD_Onda
set rds_Onda = db_Cogest.Execute(str_Sql_Onda)
if not rds_Onda.Eof then
   str_Desc_Onda = rds_Onda("ONDA_TX_DESC_ONDA")
else
   str_Desc_Onda = "Não encontrado"   
end if
rds_Onda.Close
set rds_Onda = Nothing
'==================================================================================
'==== ENCONTRA DADOS ADIOCIONAIS DA TAREFA ========================================
str_Sql_DadosAdicionais_Tarefa = ""
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " SELECT   "
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " TASK_UID"
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " , TASK_NAME"
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " , RESERVED_DATA"
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " , TASK_START_DATE"
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " , TASK_FINISH_DATE"
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " FROM MSP_TASKS"
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " WHERE PROJ_ID = " & int_Cd_ProjetoProject
str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " AND TASK_UID = " & int_Id_TarefaProject 
set rds_DadosAdicionais_Tarefa = db_Cronograma.Execute(str_Sql_DadosAdicionais_Tarefa)
if not rds_DadosAdicionais_Tarefa.Eof then
   dat_Dt_Inicio = rds_DadosAdicionais_Tarefa("TASK_START_DATE")
   dat_Dt_Termino = rds_DadosAdicionais_Tarefa("TASK_FINISH_DATE")   
   str_NomeAtividade = rds_DadosAdicionais_Tarefa("TASK_NAME")
else
   dat_Dt_Inicio = ""
   dat_Dt_Termino = ""
end if

rds_DadosAdicionais_Tarefa.close
set rds_DadosAdicionais_Tarefa = Nothing

'=======================================================================================
' ===== ENCONTRA RESPONSÁVEL PELA TAREFA ===============================================
str_Responsavel = ""
str_Responsavel = str_Responsavel & " SELECT MSP_TEXT_FIELDS.TEXT_VALUE "
str_Responsavel = str_Responsavel & " FROM MSP_TEXT_FIELDS "
str_Responsavel = str_Responsavel & " INNER JOIN MSP_CONVERSIONS ON MSP_TEXT_FIELDS.TEXT_FIELD_ID = MSP_CONVERSIONS.CONV_VALUE"
str_Responsavel = str_Responsavel & " WHERE MSP_CONVERSIONS.CONV_STRING='Task Text11'"
str_Responsavel = str_Responsavel & " AND MSP_TEXT_FIELDS.PROJ_ID = " & int_Cd_ProjetoProject
str_Responsavel = str_Responsavel & " AND MSP_TEXT_FIELDS.TEXT_REF_UID = " & int_Id_TarefaProject
set rds_Responsavel = db_Cronograma.Execute(str_Responsavel)
if not rds_Responsavel.Eof then
   str_Nome_Responsavel = rds_Responsavel("TEXT_VALUE")
else
   str_Nome_Responsavel = " não informado "   
end if

'=======================================================================================
'======== ECARREGA DADOS DOS SISTEMAS LEGADOS ==========================================
str_RespLegado = ""
str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
str_RespLegado = str_RespLegado & " USMA_CD_USUARIO "
str_RespLegado = str_RespLegado & " , USMA_TX_NOME_USUARIO "
str_RespLegado = str_RespLegado & " FROM dbo.USUARIO_MAPEAMENTO "
str_RespLegado = str_RespLegado & " Where USMA_TX_MATRICULA <> 0"
str_RespLegado = str_RespLegado & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_RespLegado = db_Cogest.Execute(str_RespLegado)
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
<link href="../../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="Head01" -->
<style type="text/css">
<!--
.bodytext {  font-family: Verdana, Arial, sans-serif; font-size: 10pt}
.textfield {  font-family: Verdana, Arial, sans-serif; font-size: 10pt}
.header {  font-family: Verdana, Arial, sans-serif; font-size: 10pt; font-weight: bold}
.footnote {  font-family: "MS Sans Serif", sans-serif; font-size: 10pt}
.code {  font-family: "Courier New", Courier, mono; font-size: 10pt; font-weight: bold}
a:hover {  color: #FF0000; text-decoration: underline overline}
a:link {  color: #0000FF; text-decoration: underline}
a:visited {  color: #0000FF; text-decoration: underline}
.highlight {  color: #FF0000; font-family: Verdana, Arial, sans-serif; font-size: 10pt}
-->
</style>
<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaScri" -->
<script language="JavaScript" src="pupdate.js"></script>

<script language="JavaScript">	
	/*
	 Nome........: VerifiCacaretersEspeciais
	 Descricao...: VERIFICA A EXITÊNCIA DE CARACTERES ESPECIAIS DURANTE A DIGITAÇÃO E OS RETIRA APÓS 
				   MSG PARA O USUÁRIO. (EVENTO - onKeyUp)
	 Paramentros.: Valor digitado pelo usuário
	 Retorno.....:
	 Autor.......: Rogério Ribeiro - DBA Engenharia de Sistemas
	 Data........: 11/06/2003
	 Obs.........:
	*/
	function VerifiCacaretersEspeciais(strvalor,strobjnome)
	{			
		var vetEspeciais = new Array();			
		var strvalor = new String(strvalor);		
					
		var i, j;
		vetEspeciais[0] = "&";
		vetEspeciais[1] = "'";
		vetEspeciais[2] = '"'
		vetEspeciais[3] = '>';
		vetEspeciais[4] = '<';			
					
		i=0;
		j=0;
					
		for (i=0; i<=strvalor.length-1; i++)
		{			
			for (j=0; j<=vetEspeciais.length-1; j++)
			{					
				if (strvalor.charAt(i) == vetEspeciais[j])
				{
					alert ('O caracter ' + strvalor.charAt(i) + ' não pode ser utilizado no texto.');
					
					if (strobjnome=='txtDescrParada')
					{
						document.forms[0].txtDescrParada.value = strvalor.substr(0,i);
					}
					else
					{
						document.forms[0].txtProcedParada.value = strvalor.substr(0,i);
					}
					break;
				}
			}
		}		
	}
		
	function confirma_ppo()
	{
		var txt_DescrParada 	= document.frm_Plano_PPO.txtDescrParada.value;			
		var int_RespTecSinGeral = document.frm_Plano_PPO.selRespTecSinGeral.selectedIndex; 
		//var int_RespFunSinGeral = document.frm_Plano_PPO.selRespFunSinGeral.selectedIndex;			
		var int_RespTecLegGeral = document.frm_Plano_PPO.selRespTecLegGeral.selectedIndex; 
		var int_RespFunLegGeral = document.frm_Plano_PPO.selRespFunLegGeral.selectedIndex; 			
		var str_TempParada 		= document.frm_Plano_PPO.txtTempParada.value; 			
		var str_ProcedParada 	= document.frm_Plano_PPO.txtProcedParada.value; 			
		var str_DtParadaLegado	= document.frm_Plano_PPO.txtDtParadaLegado.value;			
		var str_DtIniR3 		= document.frm_Plano_PPO.txtDtIniR3.value;			
		var int_UsuarioGestor 	= document.frm_Plano_PPO.selUsuarioGestor.selectedIndex;			
		var int_DtLimiteAprov 	= document.frm_Plano_PPO.txtDtLimiteAprov.value;
		
		//*** Descrição da Parada	
		if (txt_DescrParada == "")
		  {
		  alert("É obrigatório o preenchimento do campo Descrição de Parada!");
		  document.frm_Plano_PPO.txtDescrParada.focus();
		  return;
		  }
		
		//*** Responsável Técnico - Sinergia				  
		if(int_RespTecSinGeral == 0)
		  {
		  alert("É obrigatória a seleção de um Responsável Sinergia - Técnico!");
		  document.frm_Plano_PPO.selRespTecSinGeral.focus();
		  return;
		  }
		
		//*** Responsável Funcional - Sinergia 
		/*if(int_RespFunSinGeral == 0)
		  {
		  alert("É obrigatória a seleção de um Responsável Sinergia - Funcional!");
		  document.frm_Plano_PPO.selRespFunSinGeral.focus();
		  return;
		  }*/
		 
		//*** Responsável Técnico - Legado				  
		if(int_RespTecLegGeral == 0)
		  {
		  alert("É obrigatória a seleção de um Responsável Legado - Técnico!");
		  document.frm_Plano_PPO.selRespTecLegGeral.focus();
		  return;
		  } 
		 
		//*** Responsável Funcional - Legado 
		if(int_RespFunLegGeral == 0)
		  {
		  alert("É obrigatória a seleção de um Responsável Legado - Funcional!");
		  document.frm_Plano_PPO.selRespFunLegGeral.focus();
		  return;
		  } 
		 
	   //*** Tempo da Parada   	
	   if (str_TempParada == "")
		  {
		  alert("É obrigatório o preenchimento do campo Tempo da Parada!");
		  document.frm_Plano_PPO.txtTempParada.focus();
		  return;
		  }
		  
	   //*** Procedimentos da Parada   	
	   if (str_ProcedParada == "")
		  {
		  alert("É obrigatório o preenchimento do campo Procedimentos da Parada!");
		  document.frm_Plano_PPO.txtProcedParada.focus();
		  return;
		  }  			
		
	   //*** Data - Parada Legado   	
	   if (str_DtParadaLegado == "")
		  {
		  alert("É obrigatório o preenchimento do campo Data - Parada Legado!");
		  document.frm_Plano_PPO.txtDtParadaLegado.focus();
		  return;
		  } 
	
	  //*** Data - Início no R/3  	
	   if (str_DtIniR3 == "")
		  {
		  alert("É obrigatório o preenchimento do campo Data - Início no R/3!");
		  document.frm_Plano_PPO.txtDtIniR3.focus();
		  return;
		  } 
	
	   //*** Usuário Gestor	
	   if (int_UsuarioGestor == 0)
		  {
		  alert("É obrigatória a seleção de um Gestor para o processo!");
		  document.frm_Plano_PPO.selUsuarioGestor.focus();
		  return;
		  } 
		  
	   //*** Data Limite para Aprovação	
	   if (int_DtLimiteAprov == "")
		  {
		  alert("É obrigatório o preenchimento do campo Data Limite de Aprovação!");
		  document.frm_Plano_PPO.txtDtLimiteAprov.focus();
		  return;
		  } 			
	
	   document.frm_Plano_PPO.action="grava_plano.asp?pPlano=PPO";           
	   document.frm_Plano_PPO.submit();				
	}	
</script>		
<!-- InstanceEndEditable -->
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<div id="Layer1" style="position:absolute; left:20px; top:10px; width:134px; height:53px; z-index:1"><img src="../../img/000005.gif" alt=":: Logo Sinergia" width="134" height="53" border="0" usemap="#Map2"> 
	  <map name="Map2">
	    <area shape="rect" coords="6,7,129,49">
	  </map>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><table width="780" height="44" border="0" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="583" height="44"><img src="../../img/_0.gif" width="1" height="1"></td>
	          <td width="197" height="44"><img src="../../../../imagens/000043.gif" width="95" height="44"></td>
	        </tr>
	      </table></td>
	  </tr>
</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td bgcolor="#6699CC">
			<table width="780" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td width="154" height="21"><img src="../../img/000002.gif" width="154" height="21"></td>
			    <td width="19" height="21"><img src="../../img/000003.gif" width="19" height="21"></td>
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
	    <td width="1" height="1" bgcolor="#003366"><img src="../../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="5"><img src="../../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="780" height="58" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20" height="39"><img src="../../img/_0.gif" width="1" height="1"></td>
        <td width="740" height="39" background="../../img/000006.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="11%">&nbsp;</td>
              <td width="13%">&nbsp;</td>
              <td width="61%"><font color="#666666" size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>PLANO DE ENTRADA EM PRODU&Ccedil;&Atilde;O</b></font></td>
              <td width="15%"><a href="../../../../indexA_xpep.asp"><img src="../../img/botao_home_off_01.gif" alt="Ir para tela inicial" width="34" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td width="20" height="39"><img src="../../img/_0.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<!-- InstanceBeginEditable name="corpo" -->
<table width="625" border="0" align="center">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td>&nbsp;</td>
    <td><div align="center" class="campo">A&ccedil;&atilde;o</div></td>
  </tr>
  <tr>
    <td width="85"><a href="javascript:confirma_ppo()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
    <td width="25"><b></b></td>
    <td width="193"><img src="../img/limpar_01.gif" width="85" height="19"></td>
    <td width="26"></td>
    <td width="49"></td>
    <td width="27"></td>
    <td width="83">&nbsp;</td>
    <td width="103" bgcolor="#EFEFEF"><div align="center"><span class="campob"><%=str_Acao%></span></div></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="6%">&nbsp;</td>
    <td width="81%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo"><table width="75%" border="0" cellpadding="0" cellspacing="7">
        <tr> 
          <td width="11%"><div align="right" class="subtitulob">Onda:</div></td>
          <td colspan="2" class="subtitulo"><%=str_Desc_Onda%></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td width="10%" class="subtitulob">Plano:</td>
          <td class="subtitulo">Plano de Parada Operacional - PPO</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="75%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr> 
    <td width="19%" bgcolor="#EEEEEE"> <div align="right" class="campo">Atividade:</div></td>
    <td colspan="3" class="campob"><%=str_NomeAtividade%></td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Respons&aacute;vel:</div></td>
    <td colspan="3" class="campob"><%=str_Nome_Responsavel%></td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Data In&iacute;cio:</div></td>
    <td width="23%" class="campob"><%=Day(dat_Dt_Inicio) & "/" & Month(dat_Dt_Inicio) & "/" & Year(dat_Dt_Inicio) %></td>
    <td width="17%" bgcolor="#EEEEEE"> <div align="right" class="campo">Data de 
        T&eacute;rmino:</div></td>
    <td width="41%" class="campob"><%=Day(dat_Dt_Termino) & "/" & Month(dat_Dt_Termino) & "/" & Year(dat_Dt_Termino) %></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><hr></td>
  </tr>
</table>
<form name="frm_Plano_PPO" method="post" action="">
  <table width="98%" border="0">
    <tr> 
      <td colspan="5"></td>
    </tr>
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="14%" valign="top"><div align="right" class="campob">Parada:</div></td>
      <td width="44%"><textarea name="txtDescrParada" cols="50" rows="5" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"></textarea></td>
      <td width="13%">&nbsp;</td>
      <td width="27%">&nbsp;</td>
    </tr>
    <tr>
      <td height="7"></td>
      <td height="7" valign="top"></td>
      <td height="7"></td>
      <td height="7"></td>
      <td height="7"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td valign="top"><div align="right"><span class="campob">Respons&aacute;vel Sinergia:</span></div></td>
      <td><!--#include file="../includes/inc_lista_Responsavel_Sinergia_Um.asp" --></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    </tr>
    <td height="7" colspan="5"></td>
    </tr>
    <tr> 
      <td colspan="5"><!--#include file="../includes/inc_lista_Responsavel_Legado.asp" --> 
    <tr> 
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right" class="campob">Tempo da Parada:</div></td>
      <td><input name="txtTempParada" type="text" class="txtCampo" size="3">
        <select name="selUnidMedida" size="1" class="cmd150">
			<option value="Hora">Hora</option>
			<option value="Dia" selected>Dia</option>
			<option value="Mês">Mês</option>
        </select></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="7"></td>
      <td height="7" valign="top"></td>
      <td height="7"></td>
      <td height="7"></td>
      <td height="7"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td valign="top"><div align="right" class="campob">Procedimentos para a Parada:</div></td>
      <td><textarea name="txtProcedParada" cols="50" rows="5" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"></textarea></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="7" class="campo"></td>
      <td height="7" class="campo"></td>
      <td height="7" class="campo"></td>
      <td height="7"></td>
      <td height="7"></td>
    </tr>
    <tr> 
      <td class="campo">&nbsp;</td>
      <td class="campo"><div align="right" class="campob">Data da Parada do  Legado:</div></td>
      <td>        <table width="100%"  border="0">
          <tr>
            <td>            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="47%"><input name="txtDtParadaLegado" type="text" class="txtCampo" size="10" maxlength="10" readonly></td>
                <td width="53%"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0" onClick="getCalendarFor(document.body.offsetHeight,document.frm_Plano_PPO.txtDtParadaLegado)"></td>
              </tr>
            </table></td>
            <td><div align="right"><span class="campob">Início no R/3:</span></div></td>
            <td>              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="46%"><input name="txtDtIniR3" type="text" class="txtCampo" size="10" maxlength="10" readonly></td>
                  <td width="54%"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0" onClick="getCalendarFor(document.body.offsetHeight,document.frm_Plano_PPO.txtDtIniR3)"></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td height="7"></td>
      <td height="7"></td>
      <td height="7"></td>
      <td height="7"></td>
      <td height="7"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div class="campob"> 
          <div align="left">Gestor do Processo:</div>
        </div></td>
      <td><!--#include file="../includes/inc_combo_Usuario_Gestor.asp" --> </td>
      <td align="right" valign="top" class="campob"><div align="right">Data Limite 
          para aprovação:</div></td>
      <td>        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="23%"><input name="txtDtLimiteAprov" type="text" class="txtCampo" size="10" maxlength="10" onFocus="document.sample.button2.focus()" readonly></td>
            <td width="77%"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0" onClick="getCalendarFor(document.frm_Plano_PPO.txtDtLimiteAprov)"></a></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right" class="campob">Plano de Contingência:</div></td>
      <td><a href="encaminha_plano.asp?selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Cd_ProjetoProject%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
      <td class="campob"><div align="right">Plano de Comunicação:</div></td>
      <td><a href="encaminha_plano.asp?selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Cd_ProjetoProject%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right" class="campob"></div></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>
	  	<input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
		<input type="hidden" value="<%=int_Plano%>" name="pintPlano">
	  	<input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
		<input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
		<input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">
	  </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<!-- PopUp Calendar BEGIN -->
<script language="JavaScript">
if (document.all) {
 document.writeln("<div id=\"PopUpCalendar\" style=\"position:absolute; left:0px; top:0px; z-index:7; width:200px; height:77px; overflow: visible; visibility: hidden; background-color: #FFFFFF; border: 1px none #000000\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout(\'hideCalendar()\',500)\">");
 document.writeln("<div id=\"monthSelector\" style=\"position:absolute; left:0px; top:0px; z-index:9; width:181px; height:27px; overflow: visible; visibility:inherit\">");}
else if (document.layers) {
 document.writeln("<layer id=\"PopUpCalendar\" pagex=\"0\" pagey=\"0\" width=\"200\" height=\"200\" z-index=\"100\" visibility=\"hide\" bgcolor=\"#FFFFFF\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout('hideCalendar()',500)\">");
 document.writeln("<layer id=\"monthSelector\" left=\"0\" top=\"0\" width=\"181\" height=\"27\" z-index=\"9\" visibility=\"inherit\">");}
else {
 document.writeln("<p><font color=\"#FF0000\"><b>Error ! The current browser is either too old or too modern (usind DOM document structure).</b></font></p>");}
</script>
<noscript></noscript>
<table border="1" cellspacing="1" cellpadding="2" width="200" bordercolorlight="#000000" bordercolordark="#000000" vspace="0" hspace="0"><form name="ppcMonthList"><tr><td align="center" bgcolor="#CCCCCC"><a href="javascript:moveMonth('Back')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b>&lt;&nbsp;</b></font></a><font face="MS Sans Serif, sans-serif" size="1"> 
<select name="sItem" onMouseOut="if(ppcIE){window.event.cancelBubble = true;}" onChange="switchMonth(this.options[this.selectedIndex].value)" style="font-family: 'MS Sans Serif', sans-serif; font-size: 9pt"><option value="0" selected>2000
  . January</option><option value="1">2000 . February</option><option value="2">2000
  . March</option><option value="3">2000 . April</option><option value="4">2000
  . May</option><option value="5">2000 . June</option><option value="6">2000
  . July</option><option value="7">2000 . August</option><option value="8">2000
  . September</option><option value="9">2000 . October</option><option value="10">2000
  . November</option><option value="11">2000 . December</option><option value="0">2001
  . January</option></select></font><a href="javascript:moveMonth('Forward')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b>&nbsp;&gt;</b></font></a></td></tr></form></table>
<table border="1" cellspacing="1" cellpadding="2" bordercolorlight="#000000" bordercolordark="#000000" width="200" vspace="0" hspace="0"><tr align="center" bgcolor="#CCCCCC"><td width="20" bgcolor="#FFFFCC"><b><font face="MS Sans Serif, sans-serif" size="1">Su</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Mo</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Tu</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">We</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Th</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Fr</font></b></td><td width="20" bgcolor="#FFFFCC"><b><font face="MS Sans Serif, sans-serif" size="1">Sa</font></b></td></tr></table>
<script language="JavaScript">
if (document.all) {
 document.writeln("</div>");
 document.writeln("<div id=\"monthDays\" style=\"position:absolute; left:0px; top:52px; z-index:8; width:200px; height:17px; overflow: visible; visibility:inherit; background-color: #FFFFFF; border: 1px none #000000\">&nbsp;</div></div>");}
else if (document.layers) {
 document.writeln("</layer>");
 document.writeln("<layer id=\"monthDays\" left=\"0\" top=\"52\" width=\"200\" height=\"17\" z-index=\"8\" bgcolor=\"#FFFFFF\" visibility=\"inherit\">&nbsp;</layer></layer>");}
else {/*NOP*/}
</script>
<!-- PopUp Calendar END -->
<!-- InstanceEndEditable -->
    <table width="200" border="0" align="center">
<tr>	
	<td height="10" width="780"></td>
</tr>
<tr>
	<td width="780">			
		<p width="780" align="center"><img src="../../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>
