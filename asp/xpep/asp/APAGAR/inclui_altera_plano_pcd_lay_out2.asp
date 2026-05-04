<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Response.Expires=0

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

set db_Cronograma = Server.CreateObject("ADODB.Connection")
db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

str_Acao = Request("pAcao")
int_Cd_ProjetoProject = Request("pCdProjProject")

if Request("pTArefa") <> "" then
	int_Id_TarefaProject = Request("pTArefa")
else
	int_Id_TarefaProject = Request("idTaskProject")
end if

int_CD_Onda = request("pOnda")
intResData = request("pResData")

if request("pPlano") <> "" then
	int_Plano = request("pPlano")
else
	int_Plano = request("pintPlano")
end if

str_Fase = request("pFase")
strPlanoOriginal = Request("pPlanoOriginal")

str_txtRespLegadoTec 	= ""
str_txtRespLegadoFunc	= ""
str_txtRespSinergiaTec	= ""
str_txtRespSinergiaFunc	= ""
str_txtDadoMigrado 		= ""
str_DesenvAssociados	= ""
int_SistLegado			= ""
str_SistLegado			= ""
str_TipoCarga			= ""
str_TipoDados			= ""
str_CaractDado			= ""
int_txtExtracao_PCD		= ""
str_txtExtracao_Unid	= ""
int_txtCarga_PCD		= ""
str_txtCarga_Unid		= ""
str_txtArqCarga			= ""
int_txtVolume			= ""
str_txtDependencias		= ""
str_txtComoExecuta		= ""
str_txtDTExtracao_PCD 	= ""
str_txtDTCarga_PCD_Inicio = ""
str_txtDTCarga_PCD_Fim 	= ""
			
if str_Acao = "I" then
   str_Texto_Acao = "Inclusão"
else
    str_Texto_Acao = "Alteração"
   
    str_sqlGeralAlteracao = ""
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & " FROM XPEP_PLANO_TAREFA_GERAL"
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & " WHERE PLTA_NR_ID_TAREFA_PROJECT = " & int_Id_TarefaProject
	Set rds_sqlGeralAlteracao = db_Cogest.Execute(str_sqlGeralAlteracao)			
	
	if not rds_sqlGeralAlteracao.eof then		
		str_sqlAtividadeAlt = ""
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & "SELECT PLAN_NR_SEQUENCIA_PLANO"			
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PLTA_NR_SEQUENCIA_TAREFA "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_SISTEMA_LEGADO "			
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_DADO_A_SER_MIGRADO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_TIPO_ATIVIDADE "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_TIPO_DADO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_CARAC_DADO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_QTD_TEMPO_EXEC_CARGA "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_UNID_TEMPO_EXEC_CARGA "			
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_ARQ_CARGA "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_NR_VOLUME "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_DEPENDENCIAS "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_DT_EXTRACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_DT_CARGA_INICIO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_DT_CARGA_FIM "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_TX_COMO_EXECUTA "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_NR_ID_PLANO_CONTINGENCIA "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCD_NR_ID_PLANO_COMUNICACAO "				
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_LEG_TEC "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_LEG_FUN "			
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_SIN_TEC "	
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_SIN_FUN "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_TX_OPERACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_CD_NR_USUARIO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_DT_ATUALIZACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " FROM XPEP_PLANO_TAREFA_PCD"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PLTA_NR_SEQUENCIA_TAREFA = " & rds_sqlGeralAlteracao("PLTA_NR_SEQUENCIA_TAREFA")
		
		Set rds_sqlAtividadeAlt = db_Cogest.Execute(str_sqlAtividadeAlt)	
		
		if not rds_sqlAtividadeAlt.eof then		
			str_txtRespLegadoTec 	= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_LEG_TEC"))
			str_txtRespLegadoFunc	= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_LEG_FUN"))			
			str_txtRespSinergiaTec	= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_SIN_TEC"))
			str_txtRespSinergiaFunc	= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_SIN_FUN"))			
			str_txtDadoMigrado 		= Trim(rds_sqlAtividadeAlt("PPCD_TX_DADO_A_SER_MIGRADO"))			
			str_SistLegado			= rds_sqlAtividadeAlt("PPCD_TX_SISTEMA_LEGADO")			
			str_TipoCarga			= Trim(rds_sqlAtividadeAlt("PPCD_TX_TIPO_ATIVIDADE"))
			str_TipoDados			= Trim(rds_sqlAtividadeAlt("PPCD_TX_TIPO_DADO"))			
			str_CaractDado			= Trim(rds_sqlAtividadeAlt("PPCD_TX_CARAC_DADO"))
			int_txtExtracao_PCD		= rds_sqlAtividadeAlt("PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO")
			str_txtExtracao_Unid	= Trim(rds_sqlAtividadeAlt("PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO"))	
			int_txtCarga_PCD		= rds_sqlAtividadeAlt("PPCD_TX_QTD_TEMPO_EXEC_CARGA")
			str_txtCarga_Unid		= Trim(rds_sqlAtividadeAlt("PPCD_TX_UNID_TEMPO_EXEC_CARGA"))		
			str_txtArqCarga			= Trim(rds_sqlAtividadeAlt("PPCD_TX_ARQ_CARGA"))
			int_txtVolume			= rds_sqlAtividadeAlt("PPCD_NR_VOLUME")
			str_txtDependencias		= Trim(rds_sqlAtividadeAlt("PPCD_TX_DEPENDENCIAS"))					
			str_txtComoExecuta		= Trim(rds_sqlAtividadeAlt("PPCD_TX_COMO_EXECUTA"))

			strDia = ""		
			strMes = ""
			strAno = ""
			vetDTExtracao_PCD = split(Trim(rds_sqlAtividadeAlt("PPCD_DT_EXTRACAO")),"/")							
			strDia = trim(vetDTExtracao_PCD(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vetDTExtracao_PCD(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vetDTExtracao_PCD(2))
			str_txtDTExtracao_PCD = strDia & "/" & strMes & "/" & strAno 
						
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDTCargaIni = split(Trim(rds_sqlAtividadeAlt("PPCD_DT_CARGA_INICIO")),"/")							
			strDia = trim(vetDTCargaIni(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vetDTCargaIni(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vetDTCargaIni(2))
			str_txtDTCarga_PCD_Inicio = strDia & "/" & strMes & "/" & strAno 			
			
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDTCargaFim = split(Trim(rds_sqlAtividadeAlt("PPCD_DT_CARGA_FIM")),"/")							
			strDia = trim(vetDTCargaFim(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vetDTCargaFim(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vetDTCargaFim(2))
			str_txtDTCarga_PCD_Fim = strDia & "/" & strMes & "/" & strAno
		end if	
		rds_sqlAtividadeAlt.close
		set rds_sqlAtividadeAlt = nothing
	end if   
end if

'==================================================================================
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

str_RespLegado = ""
str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
str_RespLegado = str_RespLegado & " USMA_CD_USUARIO "
str_RespLegado = str_RespLegado & " , USMA_TX_NOME_USUARIO "
str_RespLegado = str_RespLegado & " FROM dbo.USUARIO_MAPEAMENTO "
str_RespLegado = str_RespLegado & " Where USMA_TX_MATRICULA <> 0"
str_RespLegado = str_RespLegado & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_RespLegado = db_Cogest.Execute(str_RespLegado)

'=======================================================================================
'======== ECARREGA DADOS DOS SISTEMAS LEGADOS ==========================================
Dim rcs_SistLegado
set rcs_SistLegado = Server.CreateObject ("ADODB.Recordset")
sql_SistLegado = ""
sql_SistLegado = sql_SistLegado & "SELECT SIST_NR_SEQUENCIAL_SISTEMA_LEGADO, SIST_TX_CD_SISTEMA, SIST_TX_DESC_SISTEMA_LEGADO"
sql_SistLegado = sql_SistLegado & " FROM XPEP_SISTEMA_LEGADO ORDER BY SIST_TX_CD_SISTEMA"
set rcs_SistLegado = db_Cogest.Execute(sql_SistLegado)
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
	<script language="javascript" src="../js/digite-cal.js"></script>	
	<script src="../js/troca_lista_sem_retirar.js" language="javascript"></script>
	<script src="../js/global.js" language="javascript"></script>	
	<script language="JavaScript">		
		function Verifica_Dif_Numeros(strValor,strNome)	
		{		
			if (isNaN(strValor))
			{				
				if (strNome == 'txtExtracao_PCD')
				{
					alert("O contéudo do campo Extração deve ser preenchido apenas com nº!");
					document.forms[0].txtExtracao_PCD.value = '';
					document.forms[0].txtExtracao_PCD.focus();
					return;
				}
				
				if (strNome == 'txtCarga_PCD')
				{
					alert("O contéudo do campo Carga deve ser preenchido apenas com nº!");
					document.forms[0].txtCarga_PCD.value = '';
					document.forms[0].txtCarga_PCD.focus();
					return;
				}
				
				if (strNome == 'txtVolume')
				{
					alert("O contéudo do campo Volume deve ser preenchido apenas com nº!");
					document.frm_Plano_PCD.txtVolume.value = '';
					document.frm_Plano_PCD.txtVolume.focus();
					return;
				}
			}
		}			
	
		/*
		 Nome........: VerifiCacaretersEspeciais
		 Descricao...: VERIFICA A EXITÊNCIA DE CARACTERES ESPECIAIS DURANTE A DIGITAÇÃO E OS RETIRA APÓS 
					   MSG PARA O USUÁRIO. (EVENTO - onKeyUp)
		 Paramentros.: Valor digitado pelo usuário
		 Retorno.....:
		 Autor.......: Rogério Ribeiro
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
						
						if (strobjnome=='txtDependencias')
						{
							document.forms[0].txtDependencias.value = strvalor.substr(0,i);
						}
						else
						{
							document.forms[0].txtComoExecuta.value = strvalor.substr(0,i);
						}
						break;
					}
				}
			}		
		}
			
		function confirma_pcd()
		{			
			var str_RespTecLegGeral = document.frm_Plano_PCD.txtRespTecLegGeral.value; 
			var str_RespFunLegGeral = document.frm_Plano_PCD.txtRespFunLegGeral.value; 	
			var str_RespTecSinGeral = document.frm_Plano_PCD.txtRespTecSinGeral.value; 
			var str_RespFunSinGeral = document.frm_Plano_PCD.txtRespFunSinGeral.value;	
			
			var int_DesenvAssoc		= document.frm_Plano_PCD.lstDesenvAssociadosSel.selectedIndex;									
			var str_DadoMigrado		= document.frm_Plano_PCD.txtDadoMigrado.value; 
			var str_SistLegado		= document.frm_Plano_PCD.txtSistLegado.value; 
			var str_Extracao_PCD	= document.frm_Plano_PCD.txtExtracao_PCD.value;	
			var str_Carga_PCD 		= document.frm_Plano_PCD.txtCarga_PCD.value;
			var str_ArqCarga		= document.frm_Plano_PCD.txtArqCarga.value;				
			var str_Volume 			= document.frm_Plano_PCD.txtVolume.value;
			var str_Dependencias	= document.frm_Plano_PCD.txtDependencias.value;
			var str_DTExtracao_PCD	= document.frm_Plano_PCD.txtDTExtracao_PCD.value;
			var srt_DTCarga_PCD_Ini	= document.frm_Plano_PCD.txtDTCarga_PCD_Inicio.value;
			var srt_DTCarga_PCD_Fim	= document.frm_Plano_PCD.txtDTCarga_PCD_Fim.value;
			var str_ComoExecuta		= document.frm_Plano_PCD.txtComoExecuta.value;
						 
			//*** Responsável Técnico - Legado				  
			if(str_RespTecLegGeral == '')
			  {
			  alert("É obrigatório o preenchimento do campo Responsável Legado - Técnico!");
			  document.frm_Plano_PCD.txtRespTecLegGeral.focus();
			  return;
			  } 
			 
			//*** Responsável Funcional - Legado 
			if(str_RespFunLegGeral == '')
			  {
			  alert("É obrigatório o preenchimento do campo Responsável Legado - Funcional!");
			  document.frm_Plano_PCD.txtRespFunLegGeral.focus();
			  return;
			  } 
			
			//*** Responsável Técnico - Sinergia				  
			if(str_RespTecSinGeral == '')
			  {
			  alert("É obrigatório o preenchimento do campo Responsável Sinergia - Técnico!");
			  document.frm_Plano_PCD.txtRespTecSinGeral.focus();
			  return;
			  }
			
			//*** Responsável Funcional - Sinergia 
			if(str_RespFunSinGeral == '')
			  {
			  alert("É obrigatória a seleção de um Responsável Sinergia - Funcional!");
			  document.frm_Plano_PCD.txtRespFunSinGeral.focus();
			  return;
			  }
			  
			//*** Dado a ser Migrado		  
			if(str_DadoMigrado == "")
			  {
			  alert("É obrigatório o preenchimento do campo Dado a ser Migrado!");
			  document.frm_Plano_PCD.txtDadoMigrado.focus();
			  return;
			  }    	
			  		  
			//*** Sistema Legado de Origem				  
			if(str_SistLegado == "")
			  {
			  alert("É obrigatório o preenchimento do campo Sistema Legado de Origem!");
			  document.frm_Plano_PCD.txtSistLegado.focus();
			  return;
			  } 			
			
		   //*** Extração  	
		   if (str_Extracao_PCD == "")
			  {
			  alert("É obrigatório o preenchimento do campo Extração(h)!");
			  document.frm_Plano_PCD.txtExtracao_PCD.focus();
			  return;
			  }			
		   else
			  {
				if (isNaN(str_Extracao_PCD))
				{
					alert("O contéudo do campo Extração(h) deve ser preenchido apenas com nº!");
					document.frm_Plano_PCD.txtExtracao_PCD.value = '';
					document.frm_Plano_PCD.txtExtracao_PCD.focus();
					return;
				}
			  }  
									  
		   //*** Carga  	
		   if (str_Carga_PCD == "")
			  {
			  alert("É obrigatório o preenchimento do campo Carga(h)!");
			  document.frm_Plano_PCD.txtCarga_PCD.focus();
			  return;
			  }  
			else
			  {
				if (isNaN(str_Carga_PCD))
				{
					alert("O contéudo do campo Carga(h) deve ser preenchido apenas com nº!");
					document.frm_Plano_PCD.txtCarga_PCD.value = '';
					document.frm_Plano_PCD.txtCarga_PCD.focus();
					return;
				}
			  }  
			    
		   //*** Arquivos de carga   	
		   if (str_ArqCarga == "")
			  {
			  alert("É obrigatório o preenchimento do campo Arquivos de Carga!");
			  document.frm_Plano_PCD.txtArqCarga.focus();
			  return;
			  }	
		
		  //*** Volume
		   if (str_Volume == "")
			  {
			  alert("É obrigatório o preenchimento do campo Volume!");
			  document.frm_Plano_PCD.txtVolume.focus();
			  return;
			  }		  	
			
		   //*** Dependências
		   if (str_Dependencias == "")
			  {
			  alert("É obrigatório o preenchimento do campo Dependências!");
			  document.frm_Plano_PCD.txtDependencias.focus();
			  return;
			  } 	
				
		   //*** Data Extração
		   if (str_DTExtracao_PCD == "")
			  {
			  alert("É obrigatório o preenchimento do campo Data Extração!");
			  document.frm_Plano_PCD.txtDTExtracao_PCD.focus();
			  return;
			  }
		   else
		     {
			 	validaData(str_DTExtracao_PCD,'txtDTExtracao_PCD','Data Extração');
				if (blnData) return; 	
		     }
		
		   //*** Data Carga - Inicio
		   if (srt_DTCarga_PCD_Ini == "")
			  {
			  alert("É obrigatório o preenchimento do campo Data Carga - Início!");
			  document.frm_Plano_PCD.txtDTCarga_PCD_Inicio.focus();
			  return;
			  }
		   else
		     {
			 	validaData(srt_DTCarga_PCD_Ini,'txtDTCarga_PCD_Inicio','Data Carga - Inicio');
				if (blnData) return; 	
		     } 	
				
		   //*** Data Carga - Fim
		   if (srt_DTCarga_PCD_Fim == "")
			  {
			  alert("É obrigatório o preenchimento do campo Data Carga - Fim!");
			  document.frm_Plano_PCD.txtDTCarga_PCD_Fim.focus();
			  return;
			  }
		   else
		     {
			 	validaData(srt_DTCarga_PCD_Fim,'txtDTCarga_PCD_Fim','Data Carga - Fim');
				if (blnData) return; 	
		     } 	 		
				
		   //*** Como executa
		   if (str_ComoExecuta == "")
			  {
			  alert("É obrigatório o preenchimento do campo Como Executa!");
			  document.frm_Plano_PCD.txtComoExecuta.focus();
			  return;
			  } 				
			
		   carrega_txt(document.frm_Plano_PCD.lstDesenvAssociadosSel)	
			
		   function carrega_txt(fbox) 
			{			
				document.frm_Plano_PCD.pSistemas.value = "";
				for(var i=0; i<fbox.options.length; i++) 
				{				
					document.frm_Plano_PCD.pSistemas.value = document.frm_Plano_PCD.pSistemas.value + "|" + fbox.options[i].value;						
				}	
			}			
		   document.frm_Plano_PCD.action="grava_plano.asp?pPlano=PCD";           
		   document.frm_Plano_PCD.submit();				
		}	
		
		function Localiza_Usuario(strTipoResponsavel,strCampo)
		{	
			if (strCampo == 'txtRespTecLegGeral')
			{
				strUsuario = document.frm_Plano_PCD.txtRespTecLegGeral.value;
				
				if (strUsuario == '')
				{			
					alert("É obrigatório o preenchimento do campo Responsável Legado - Técnico!");
					document.frm_Plano_PCD.txtAprovadorPB.focus();
					return;
				}		
			}
				
			if (strCampo == 'txtRespFunLegGeral')
			{
				strUsuario = document.frm_Plano_PCD.txtRespFunLegGeral.value;
				
				if (strUsuario == '')
				{			
					alert("É obrigatório o preenchimento do campo Responsável Legado - Funcional!");
					document.frm_Plano_PCD.txtRespFunLegGeral.focus();
					return;
				}		
			}						
					
			if (strCampo == 'txtRespTecSinGeral')
			{
				strUsuario = document.frm_Plano_PCD.txtRespTecSinGeral.value;
				
				if (strUsuario == '')
				{			
					alert("É obrigatório o preenchimento do campo Responsável Sinergia - Técnico!");
					document.frm_Plano_PCD.txtRespTecSinGeral.focus();
					return;
				}		
			}					
			
			if (strCampo == 'txtRespFunSinGeral')
			{
				strUsuario = document.frm_Plano_PCD.txtRespFunSinGeral.value;
				
				if (strUsuario == '')
				{			
					alert("É obrigatório o preenchimento do campo Responsável Sinergia - Funcional!");
					document.frm_Plano_PCD.txtRespFunSinGeral.focus();
					return;
				}		
			}				
					
			document.frm_Plano_PCD.pTipoResponsavel.value = strTipoResponsavel;	
			document.frm_Plano_PCD.pChaveUsua.value = strUsuario.toUpperCase();	
			document.frm_Plano_PCD.pCampo.value = strCampo;
							
			document.frm_Plano_PCD.action='inclui_altera_plano_pcd.asp?pTipoResponsavel=' + strTipoResponsavel + '&pCampo=' + strCampo;
			document.frm_Plano_PCD.submit();			
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
    <td width="85">&nbsp;</td>
    <td width="25"><b></b></td>
    <td width="193"><!--<img src="../img/limpar_01.gif" width="85" height="19">--></td>
    <td width="26"></td>
    <td width="49"></td>
    <td width="27"></td>
    <td width="83">&nbsp;</td>
    <td width="103" bgcolor="#EFEFEF"><div align="center"><span class="campob"><%=str_Texto_Acao%></span></div></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="6%">&nbsp;</td>
    <td width="81%" class="subtitulo"><table width="75%" border="0" cellpadding="0" cellspacing="7">
        <tr>
          <td>&nbsp;</td>
          <td colspan="2" class="subtitulo">&nbsp;</td>
        </tr>
        <tr> 
          <td width="11%"><div align="right" class="subtitulob">Onda:</div></td>
          <td colspan="2" class="subtitulo"><%=str_Desc_Onda%></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td width="10%" class="subtitulob">Plano:</td>
          <td class="subtitulo">Plano de Convers&otilde;es de Dados - PCD</td>
        </tr>
      </table></td>
    <td width="13%">&nbsp;</td>
  </tr>
</table>
<table width="75%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr> 
    <td width="17%" bgcolor="#EEEEEE"> <div align="right" class="campo">Atividade:</div></td>
    <td colspan="3" class="campob"><%=str_NomeAtividade%></td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"><div align="right" class="campo">Respons&aacute;vel:</div></td>
    <td colspan="3" class="campob"><%=str_Nome_Responsavel%></td>
  </tr>
  
  <%
  '*** DATA INÍCIO
  if cint(Day(dat_Dt_Inicio)) < 10 then 
  	strDiaInicio = "0" & Day(dat_Dt_Inicio)
  else
  	strDiaInicio = Day(dat_Dt_Inicio)
  end if
  
  if cint(Month(dat_Dt_Inicio)) < 10 then 
  	strMesInicio = "0" & Month(dat_Dt_Inicio)
  else
  	strMesInicio = Month(dat_Dt_Inicio)
  end if
  dat_Dt_Inicio = strDiaInicio & "/" & strMesInicio & "/" & Year(dat_Dt_Inicio)
  
  '*** DATA FIM
  if cint(Day(dat_Dt_Termino)) < 10 then 
  	strDiaFim = "0" & Day(dat_Dt_Termino)
  else
  	strDiaFim = Day(dat_Dt_Termino)
  end if
  
  if cint(Month(dat_Dt_Termino)) < 10 then 
  	strMesFim = "0" & Month(dat_Dt_Termino)
  else
  	strMesFim = Month(dat_Dt_Termino)
  end if
  dat_Dt_Termino = strDiaFim & "/" & strMesFim & "/" & Year(dat_Dt_Termino)
  %>
  
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Data In&iacute;cio:</div></td>
    <td width="21%" class="campob"><%=dat_Dt_Inicio%></td>
    <td width="20%" bgcolor="#EEEEEE"><div align="right" class="campo">Data de T&eacute;rmino:</div></td>
    <td width="33%" class="campob"><%=dat_Dt_Termino%></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td height="2" bgcolor="#CCCCCC"></td>
  </tr>
</table>
<form name="frm_Plano_PCD" method="post" action="">
  <table width="98%" border="0" cellpadding="2" cellspacing="2">
     <tr> 
      <td colspan="5">
	  	<table width="88%" border="0">
		  <tr> 
			<td height="27" colspan="3"> <table width="50%" border="0">
				<tr> 
				  <td width="3%">&nbsp;</td>
				  <td width="97%" class="campob">Respons&aacute;vel Legado</td>
				</tr>
			  </table></td>
			<td width="29%">&nbsp;</td>
			<td width="29%">&nbsp;</td>
		  </tr>
		  
		  <%
			Public Function RetornaNomeUsuario(strChave, strTipoResponsavel)		
				sql_VerUsuario= ""				
				sql_VerUsuario = sql_VerUsuario & " SELECT USUA_TX_NOME_USUARIO"		
				sql_VerUsuario = sql_VerUsuario & " FROM XPEP_EQUIPE_SINERGIA "
				sql_VerUsuario = sql_VerUsuario & " WHERE USUA_TX_CD_USUARIO = '" & strChave & "'"					
								
				set rds_VerUsuario = db_Cogest.Execute(sql_VerUsuario)
				
				if not rds_VerUsuario.eof then									
					RetornaNomeUsuario = Ucase(rds_VerUsuario("USUA_TX_NOME_USUARIO"))					
				else
					sql_VerUsuarioLegado = ""
					sql_VerUsuarioLegado = sql_VerUsuarioLegado & " SELECT USMA_TX_NOME_USUARIO"		
					sql_VerUsuarioLegado = sql_VerUsuarioLegado & " FROM USUARIO_MAPEAMENTO "
					sql_VerUsuarioLegado = sql_VerUsuarioLegado & " WHERE USMA_TX_MATRICULA <> 0"
					sql_VerUsuarioLegado = sql_VerUsuarioLegado & " AND USMA_CD_USUARIO = '" & strChave & "'"					
					set rds_VerUsuarioLegado = db_Cogest.Execute(sql_VerUsuarioLegado)
					
					if not rds_VerUsuarioLegado.eof then
						RetornaNomeUsuario = Ucase(rds_VerUsuarioLegado("USMA_TX_NOME_USUARIO"))
					else
						RetornaNomeUsuario = "USUÁRIO NÃO LOCALIZADO."
					end if
					rds_VerUsuarioLegado.close
					set rds_VerUsuarioLegado = nothing
				end if		
				rds_VerUsuario.close
				set rds_VerUsuario = nothing
			End function
			%>
			
			<%
			strChaveUsuario 	= Request("pChaveUsua")
			strTipoResponsavel 	= Request("pTipoResponsavel")	
			strCampo 			= Request("pCampo")			
			if strTipoResponsavel <> "" then				
				if strCampo = "txtRespTecLegGeral" then
					strUsuaRespTecLegado 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 	
				elseif strCampo = "txtRespFunLegGeral" then
					strUsuaRespFuncLegado 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 
				elseif strCampo = "txtRespTecSinGeral" then
					strUsuaRespTecSinergia 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 	
				elseif strCampo = "txtRespFunSinGeral" then
					strUsuaRespFuncSinergia 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 
				end if								
			end if
			%>
			
		  	<%if strUsuaRespTecLegado <> "" then%>	
				<input type="hidden" value="<%=strUsuaRespTecLegado%>" name="hdUsuaRespTecLegado">
			<%else%>
				<input type="hidden" value="<%=Request("hdUsuaRespTecLegado")%>" name="hdUsuaRespTecLegado">
			<%end if%>		
			
			<%if strUsuaRespFuncLegado <> "" then%>	
				<input type="hidden" value="<%=strUsuaRespFuncLegado%>" name="hdUsuaRespFuncLegado">
			<%else%>
				<input type="hidden" value="<%=Request("hdUsuaRespFuncLegado")%>" name="hdUsuaRespFuncLegado">
			<%end if%>		
			
			<%if strUsuaRespTecSinergia <> "" then%>	
				<input type="hidden" value="<%=strUsuaRespTecSinergia%>" name="hdUsuaRespTecSinergia">
			<%else%>
				<input type="hidden" value="<%=Request("hdUsuaRespTecSinergia")%>" name="hdUsuaRespTecSinergia">
			<%end if%>	
			
			<%if strUsuaRespFuncSinergia <> "" then%>	
				<input type="hidden" value="<%=strUsuaRespFuncSinergia%>" name="hdUsuaRespFuncSinergia">
			<%else%>
				<input type="hidden" value="<%=Request("hdUsuaRespFuncSinergia")%>" name="hdUsuaRespFuncSinergia">
			<%end if%>
			  
		  <tr> 
			<td width="1%" valign="top" class="campo">&nbsp;</td>
			<td width="17%" valign="top"> <div align="right"> 
				 <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr> 
					<td height="25"> <div align="right"><span class="campob">T&eacute;cnico:</span>:</div></td>
				  </tr>
				</table>
			  </div>			  
		    </td>
			<td colspan="3" class="campob">
				<%if Request("txtRespTecLegGeral") <> "" then%>
					<input type="text" maxlength="4" value="<%=Request("txtRespTecLegGeral")%>" name="txtRespTecLegGeral" size="5">
				<%else%>	  	
					<input type="text" maxlength="4" value="<%=str_txtRespLegadoTec%>" name="txtRespTecLegGeral" size="5">
				<%end if%>
				<a href="javascript:Localiza_Usuario('Legado','txtRespTecLegGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
				<%
				if strUsuaRespTecLegado <> "" then
					Response.write strUsuaRespTecLegado
				else
					Response.write Request("hdUsuaRespTecLegado") 
				end if
				%>	
			</td>			
		  </tr>
		  <tr> 
			<td valign="top">&nbsp;</td>
			<td valign="top"> <div align="right"> 
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr> 
					<td height="25"> <div align="right"><span class="campob">Funcional:</span>: 
					  </div></td>
				  </tr>
				</table>
			  </div></td>
			<td colspan="3" class="campob">
				<%if Request("txtRespFunLegGeral") <> "" then%>
					<input type="text" maxlength="4" value="<%=Request("txtRespFunLegGeral")%>" name="txtRespFunLegGeral" size="5">
				<%else%>	  	
					<input type="text" maxlength="4" value="<%=str_txtRespLegadoFunc%>" name="txtRespFunLegGeral" size="5">
				<%end if%>
				<a href="javascript:Localiza_Usuario('Legado','txtRespFunLegGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
				<%
				if strUsuaRespFuncLegado <> "" then
					Response.write strUsuaRespFuncLegado
				else
					Response.write Request("hdUsuaRespFuncLegado") 
				end if
				%>				
			</td>
		  </tr>
		</table>	  	
	  </td>
    <tr> 
	<tr> 
      <td colspan="5">
	  	<table width="88%" border="0">
		  <tr> 
			<td height="27" colspan="3"> <table width="50%" border="0">
				<tr> 
				  <td width="3%">&nbsp;</td>
				  <td width="97%" class="campob">Respons&aacute;vel Sinergia</td>
				</tr>
			  </table></td>
			<td width="29%">&nbsp;</td>
			<td width="29%">&nbsp;</td>
		  </tr>
		  <tr> 
			<td width="1%" valign="top" class="campo">&nbsp;</td>
			<td width="17%" valign="top"> <div align="right"> 
				 <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr> 
					<td height="25"> <div align="right"><span class="campob">T&eacute;cnico:</span>:</div></td>
				  </tr>
				</table>
			  </div></td>
			<td colspan="3" class="campob">
				<%if Request("txtRespTecSinGeral") <> "" then%>
					<input type="text" maxlength="4" value="<%=Request("txtRespTecSinGeral")%>" name="txtRespTecSinGeral" size="5">
				<%else%>	  	
					<input type="text" maxlength="4" value="<%=str_txtRespSinergiaTec%>" name="txtRespTecSinGeral" size="5">
				<%end if%>
				<a href="javascript:Localiza_Usuario('Sinergia','txtRespTecSinGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
				<%
				if strUsuaRespTecSinergia <> "" then
					Response.write strUsuaRespTecSinergia
				else
					Response.write Request("hdUsuaRespTecSinergia") 
				end if
				%>				
			</td>
		  </tr>
		  <tr> 
			<td valign="top">&nbsp;</td>
			<td valign="top"> <div align="right"> 
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr> 
					<td height="25"> <div align="right"><span class="campob">Funcional:</span>: 
					  </div></td>
				  </tr>
				</table>
			  </div></td>
			<td colspan="3" class="campob">
				<%if Request("txtRespFunSinGeral") <> "" then%>
					<input type="text" maxlength="4" value="<%=Request("txtRespFunSinGeral")%>" name="txtRespFunSinGeral" size="5">
				<%else%>	  	
					<input type="text" maxlength="4" value="<%=str_txtRespSinergiaFunc%>" name="txtRespFunSinGeral" size="5">
				<%end if%>
				<a href="javascript:Localiza_Usuario('Sinergia','txtRespFunSinGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
				<%
				if strUsuaRespFuncSinergia <> "" then
					Response.write strUsuaRespFuncSinergia
				else
					Response.write Request("hdUsuaRespFuncSinergia") 
				end if
				%>				
			</td>
		  </tr>
		</table>	  	
		<!--#include file="../includes/inc_lista_desenvolvimentos.asp" -->
	  </td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td height="25" class="campob" align="right">Dado a ser Migrado: </td>
      <td>
	  	<%if Request("txtDadoMigrado") <> "" then%>
			<input name="txtDadoMigrado" size="45" type="text" class="txtCampo" value="<%=Request("txtDadoMigrado")%>">
		<%else%>	  	
			<input name="txtDadoMigrado" size="45" type="text" class="txtCampo" value="<%=str_txtDadoMigrado%>">
		<%end if%>	  	
	  </td>
      <td class="campo">&nbsp;</td>
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
      <td class="campo">&nbsp;</td>
      <td height="25" class="campob"><div align="right">Sistema Legados de Origem:</div></td>
      <td>
	  	<%if Request("txtSistLegado") <> "" then%>
			<input name="txtSistLegado" size="45" type="text" class="txtCampo" value="<%=Request("txtSistLegado")%>">
		<%else%>	  	
			<input name="txtSistLegado" size="45" type="text" class="txtCampo" value="<%=str_SistLegado%>">
		<%end if%>	  
	  
		  <!--
			<select name="selSistLegado">
			  <option value="1">== Selecione um Sistema ==</option>
			  <%
				'rcs_SistLegado.MoveFirst
				'do while not rcs_SistLegado.eof
					'if int_SistLegado = cint(rcs_SistLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO")) then%>
						<option value="<%'=rcs_SistLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO")%>" selected><%'=rcs_SistLegado("SIST_TX_CD_SISTEMA") & " - " & rcs_SistLegado("SIST_TX_DESC_SISTEMA_LEGADO")%></option>
					<%'else%>
						<option value="<%'=rcs_SistLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO")%>"><%'=rcs_SistLegado("SIST_TX_CD_SISTEMA") & " - " & rcs_SistLegado("SIST_TX_DESC_SISTEMA_LEGADO")%></option>
					<%'end if
					'rcs_SistLegado.MoveNext
				'loop			 		 
				'rcs_SistLegado.close()
				'set rcs_SistLegado = nothing
				%>
			</select>
			-->
      </td>
      <td class="campob"><div align="right">Tipo de Ativ.para carga:</div></td>
      <td>
        <select name="selTipoCarga" class="cmdOnda">
          <%if str_TipoCarga = "Manual" then%>
		  	<option value="Manual" selected>Manual</option>
		  <%elseif Request("selTipoCarga") = "Manual" then%>
		  	<option value="Manual" selected>Manual</option>
		  <%else%>
		  	 <option value="Manual">Manual</option>
		  <%end if%>
		  
		  <%if str_TipoCarga = "Automática" then%>
		  	<option value="Automática" selected>Automática</option>
		  <%elseif Request("selTipoCarga") = "Automática" then%>
		  	<option value="Automática" selected>Automática</option>
		  <%else%>
		  	 <option value="Automática">Automática</option>
		  <%end if%>
		  
		  <%if str_TipoCarga = "Customizada" then%>
		  	<option value="Customizada" selected>Customizada</option>
		  <%elseif Request("selTipoCarga") = "Customizada" then%>
		  	<option value="Customizada" selected>Customizada</option>
		  <%else%>
		  	 <option value="Customizada">Customizada</option>
		  <%end if%>
         
          <%if str_TipoCarga = "Verificação" then%>
		  	<option value="Verificação" selected>Verificação</option>
		  <%elseif Request("selTipoCarga") = "Verificação" then%>
		  	<option value="Verificação" selected>Verificação</option>
		  <%else%>
		  	 <option value="Verificação">Verificação</option>
		  <%end if%>          
        </select>
      </td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td height="25" class="campob"><div align="right">Tipo de dado:</div></td>
      <td class="campo">
        <select name="selTipoDados" class="cmd150">
          	<%if str_TipoDados = "Texto" then%>
		  		<option value="Texto" selected>Texto</option>
			<%elseif Request("selTipoDados") = "Texto" then%>
		  		<option value="Texto" selected>Texto</option>
			<%else%>
				<option value="Texto">Texto</option>
			<%end if%>
			
			<%if str_TipoDados = "Idoc" then%>
		  		<option value="Idoc" selected>Idoc</option>
			<%elseif Request("selTipoDados") = "Idoc" then%>
		  		<option value="Idoc" selected>Idoc</option>
			<%else%>
				<option value="Idoc">Idoc</option>
			<%end if%>  
			
			<%if str_TipoDados = "NA" then%>
		  		<option value="NA" selected>NA</option>
			<%elseif Request("selTipoDados") = "NA" then%>
		  		<option value="NA" selected>NA</option>
			<%else%>
				<option value="NA">NA</option>
			<%end if%>        
        </select>
      </td>
      <td class="campob"><div align="right">Caracter&iacute;stica do dado:</div></td>
      <td>
        <select name="selCaractDado">
          <%if str_CaractDado  = "Mestre" then%>
		  	<option value="Mestre" selected>Mestre</option>
		  <%elseif Request("selCaractDado") = "Mestre" then%>
		  	<option value="Mestre" selected>Mestre</option>
		  <%else%>
		   	<option value="Mestre">Mestre</option>
		  <%end if%>
		  
		  <%if str_CaractDado  = "Transacional" then%>
		  	<option value="Transacional" selected>Transacional</option>
		  <%elseif Request("selCaractDado") = "Transacional" then%>
		  	<option value="Transacional" selected>Transacional</option>
		  <%else%>
		   	<option value="Transacional">Transacional</option>
		  <%end if%>          
        </select>
      </td>
    </tr>
    <tr>
      <td height="33" colspan="3">
        <table width="44%" border="0">
          <tr>
            <td width="9%">&nbsp;</td>
            <td width="91%" class="campob">Tempo de Execu&ccedil;&atilde;o</td>
          </tr>
      </table></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td>
        <div align="right" class="campob">Extra&ccedil;&atilde;o:</div></td>
      <td>
	  	<%if Request("txtExtracao_PCD") <> "" then%>
	  		<input name="txtExtracao_PCD" type="text" size="3" class="txtCampo" value="<%=Request("txtExtracao_PCD")%>" onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);">		
		<%else%>
			<input name="txtExtracao_PCD" type="text" size="3" class="txtCampo" value="<%=int_txtExtracao_PCD%>" onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);">		
		<%end if%>
		
		<select name="txtExtracao_Unid" size="1" class="cmd150">          
		  <%if str_txtExtracao_Unid = "Hora" then%>
			<option value="Hora" selected>Hora</option>
		  <%elseif Request("txtExtracao_Unid") = "Hora" then%>
		  	<option value="Hora" selected>Hora</option>
		  <%else%>
			<option value="Hora">Hora</option>
		  <%end if%>
		  
		  <%if str_txtExtracao_Unid = "Dia Útil" then%>
			<option value="Dia Útil" selected>Dia Útil</option>
		  <%elseif Request("txtExtracao_Unid") = "Dia Útil" then%>
		  	<option value="Dia Útil" selected>Dia Útil</option>
		  <%else%>
			<option value="Dia Útil">Dia Útil</option>
		  <%end if%>
		  
		  <%if str_txtExtracao_Unid = "Dia Corrido" then%>
			<option value="Dia Corrido" selected>Dia Corrido</option>
		  <%elseif Request("txtExtracao_Unid") = "Dia Corrido" then%>
		  	<option value="Dia Corrido" selected>Dia Corrido</option>
		  <%else%>
			<option value="Dia Corrido">Dia Corrido</option>
		  <%end if%>
		  
		  <%if str_txtExtracao_Unid = "Mês" then%>
			<option value="Mês" selected>Mês</option>
		  <%elseif Request("txtExtracao_Unid") = "Mês" then%>
		  	<option value="Mês" selected>Mês</option>
		  <%else%>
			<option value="Mês">Mês</option>
		  <%end if%> 
		</select>	
	  </td>
      <td><div align="right" class="campob">Carga:</div></td>
      <td>
	  	<%if Request("txtCarga_PCD") <> "" then%>
	  		<input name="txtCarga_PCD" type="text" size="3" class="txtCampo" value="<%=Request("txtCarga_PCD")%>" onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);">
		<%else%>
			<input name="txtCarga_PCD" type="text" size="3" class="txtCampo" value="<%=int_txtCarga_PCD%>" onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);">
		<%end if%>
	  		  	
		<select name="txtCarga_Unid" size="1" class="cmd150">          
		  <%if str_txtCarga_Unid = "Hora" then%>
			<option value="Hora" selected>Hora</option>
		  <%elseif Request("txtCarga_Unid") = "Hora" then%>
		  	<option value="Hora" selected>Hora</option>
		  <%else%>
			<option value="Hora">Hora</option>
		  <%end if%>
		  
		  <%if str_txtCarga_Unid = "Dia Útil" then%>
			<option value="Dia Útil" selected>Dia Útil</option>
		  <%elseif Request("txtCarga_Unid") = "Dia Útil" then%>
		  	<option value="Dia Útil" selected>Dia Útil</option>
		  <%else%>
			<option value="Dia Útil">Dia Útil</option>
		  <%end if%>
		  
		  <%if str_txtCarga_Unid = "Dia Corrido" then%>
			<option value="Dia Corrido" selected>Dia Corrido</option>
		  <%elseif Request("txtCarga_Unid") = "Dia Corrido" then%>
		  	<option value="Dia Corrido" selected>Dia Corrido</option>
		  <%else%>
			<option value="Dia Corrido">Dia Corrido</option>
		  <%end if%>
		  
		   <%if str_txtCarga_Unid = "Mês" then%>
			<option value="Mês" selected>Mês</option>
		  <%elseif Request("txtCarga_Unid") = "Mês" then%>
		  	<option value="Mês" selected>Mês</option>
		  <%else%>
			<option value="Mês">Mês</option>
		  <%end if%> 
		</select>	
	  </td>
    </tr>
      
    <tr> 
      <td colspan="5"><hr></td>
    </tr>      
        
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%" valign="top"><div align="right" class="campob">Arquivos de Carga:</div></td>
      <td width="18%">
	  	<%if Request("txtArqCarga") <> "" then%>
	  		<textarea name="txtArqCarga" cols="45" rows="4" class="txtCampo" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=Request("txtArqCarga")%></textarea></td>
		<%else%>
			<textarea name="txtArqCarga" cols="45" rows="4" class="txtCampo"><%=str_txtArqCarga%></textarea></td>
		<%end if%>	  	
      <td width="19%" valign="top"><div class="campob" align="right">Volume:</div></td>
      <td width="44%" valign="top">
	  	<%if Request("txtVolume") <> "" then%>
	  		<input name="txtVolume" type="text" class="txtCampo" size="15" maxlength="10" value="<%=Request("txtVolume")%>" onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);"></td>
		<%else%>
			<input name="txtVolume" type="text" class="txtCampo" size="15" maxlength="10" value="<%=int_txtVolume%>" onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);"></td>
		<%end if%>	  	
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td valign="top"><div align="right" class="campob">Dependências:</div></td>
      <td>
	  <%if Request("txtDependencias") <> "" then%>
		<textarea name="txtDependencias" cols="34" rows="4" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=Request("txtDependencias")%></textarea></td>
	  <%else%>
		<textarea name="txtDependencias" cols="34" rows="4" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtDependencias%></textarea></td>
	  <%end if%>		
	  
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
      
    <tr>
	  <td height="30">&nbsp;</td>
	  <td height="30"><div align="right" class="campob">Data Extração:</div></td>
      <td height="30">
	  	<%if Request("txtDTExtracao_PCD") <> "" then%>
			<input name="txtDTExtracao_PCD" size="10" class="txtCampo" type="text" readonly value="<%=Request("txtDTExtracao_PCD")%>">
	  	<%else%>
			<input name="txtDTExtracao_PCD" size="10" class="txtCampo" type="text" readonly value="<%=str_txtDTExtracao_PCD%>">
	  	<%end if%>		
	  	
	  	<a href="javascript:show_calendar(true,'frm_Plano_PCD.txtDTExtracao_PCD','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a>
	</td>
      <td height="30">&nbsp;</td>
      <td height="30">&nbsp;</td>
    </tr>
    
    <tr> 	  
      <td colspan="5"><table width="100%"  border="0">
        <tr>
          <td width="2%">&nbsp;</td>
          <td width="14%" align="right" class="campob"> Data Carga in&iacute;cio:</td>
          <td width="19%">
		  	<%if Request("txtDTCarga_PCD_Inicio") <> "" then%>
				<input name="txtDTCarga_PCD_Inicio" type="text" size="10" class="txtCampo" readonly value="<%=Request("txtDTCarga_PCD_Inicio")%>">
			<%else%>
				<input name="txtDTCarga_PCD_Inicio" type="text" size="10" class="txtCampo" readonly value="<%=str_txtDTCarga_PCD_Inicio%>">
			<%end if%>			  	
            <a href="javascript:show_calendar(true,'frm_Plano_PCD.txtDTCarga_PCD_Inicio','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a> </td>
          <td width="25%" align="right" class="campob"> Data Carga fim:</td>
          <td width="40%">
		  	<%if Request("txtDTCarga_PCD_Fim") <> "" then%>
				<input name="txtDTCarga_PCD_Fim" type="text" size="10" class="txtCampo" readonly value="<%=Request("txtDTCarga_PCD_Fim")%>">
			<%else%>
				<input name="txtDTCarga_PCD_Fim" type="text" size="10" class="txtCampo" readonly value="<%=str_txtDTCarga_PCD_Fim%>">
			<%end if%>		  
            <a href="javascript:show_calendar(true,'frm_Plano_PCD.txtDTCarga_PCD_Fim','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a> </td>
        </tr>
      </table></td>
    </tr>
	
    <tr>
      <td>&nbsp;</td>
      <td valign="top"><div align="right" class="campob">Como Executa:</div></td>
      <td>
	  	<%if Request("txtComoExecuta") <> "" then%>
			<textarea name="txtComoExecuta" cols="34" rows="4" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=Request("txtComoExecuta")%></textarea></td>
		<%else%>
			<textarea name="txtComoExecuta" cols="34" rows="4" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtComoExecuta%></textarea></td>
		<%end if%>	
	  	
      <td>
	  	<input type="hidden" value="<%=int_Cd_ProjetoProject%>" name="pCdProjProject">
	    <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
        <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
        <input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
        <input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
        <input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">
        <input type="hidden" value="<%=str_Acao%>" name="pAcao">
        <input type="hidden" name="pSistemas" value="">
		<input type="hidden" value="<%=int_CD_Onda%>" name="pOnda">
	    <input type="hidden" value="<%=str_Fase%>" name="pFase">
	    <input type="hidden" value="<%=strPlanoOriginal%>" name="pPlanoOriginal">	
		<input type="hidden" value="" name="pCampo">
		<input type="hidden" value="" name="pTipoResponsavel">
		<input type="hidden" value="" name="pChaveUsua">		
	  </td>
      <td>&nbsp;</td>
    </tr>
	
	<%if str_Acao = "A" then%>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right"><span class="campob">Plano de Contingência:</span>  
      </div></td>
      <td><a href="encaminha_plano.asp?selTipoCadastro=PAC&pSiglaPlano=PAC&pAtividade_Origen=<%="PCD - " & str_NomeAtividade%>&selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Plano%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	<%end if%>
  </table>
  <table width="625" border="0" align="center">
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>&nbsp;</td>
      <td><div align="center" class="campo"></div></td>
    </tr>
    <tr>
      <td width="85"><a href="javascript:confirma_pcd()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
      <td width="25"><b></b></td>
      <td width="193"><!--<img src="../img/limpar_01.gif" width="85" height="19">--></td>
      <td width="26"></td>
      <td width="49"></td>
      <td width="27"></td>
      <td width="83">&nbsp;</td>
      <td width="103"><div align="center"></div></td>
    </tr>
  </table>
</form>
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
