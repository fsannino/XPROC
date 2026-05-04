<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Response.Expires=0

on error resume next
	set db_Cogest = Server.CreateObject("ADODB.Connection")
	db_Cogest.Open Session("Conn_String_Cogest_Gravacao")
	
	set db_Cronograma = Server.CreateObject("ADODB.Connection")
	db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

if err.number <> 0 then		
	strMSG = "Ocorreu algum problema com o servidor!"
	Response.redirect "msg_geral.asp?pMsg=" & strMsg & "&pErroServidor=S"
end if	
		
str_Acao = Request("pAcao")

int_Cd_ProjetoProject = Request("pCdProjProject")

if Request("pTArefa") <> "" then
	int_Id_TarefaProject = Request("pTArefa")
else
	int_Id_TarefaProject = Request("idTaskProject")
end if

int_CD_Onda = request("pOnda")
int_ResData = request("pResData")

if request("pPlano") <> "" then 
	int_Plano = request("pPlano")
else
	int_Plano = request("pintPlano")
end if

if request("pIdAtividade") <> "" then 
	int_IdAtividade = request("pIdAtividade")
end if

'Response.write int_Id_TarefaProject & "<br>"
'Response.write int_IdAtividade 
'Response.end

str_txtDescrParada 		= ""
str_txtRespSinergia		= ""
str_txtRespLegadoTec	= ""
str_txtRespLegadoFunc	= ""
str_txtTempParada  		= ""
str_txtUnidTempo		= ""
str_txtProcedParada		= ""
str_txtDtParadaLegado	= ""
str_txtDtIniR3		 	= ""
str_txtDtLimiteAprov	= ""
str_UsuarioGestor		= ""

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
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PLAN_NR_SEQUENCIA_PLANO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PLTA_NR_SEQUENCIA_TAREFA"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_TX_DESCRICAO_PARADA"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_TX_QTD_TEMPO_PARADA"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_TX_UNID_TEMPO_PARADA"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_TX_PROCEDIMENTOS_PARADA"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_DT_PARADA_LEGADO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_DT_INICIO_R3"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_DT_LIMITE_APROVACAO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_NR_ID_PLANO_CONTINGENCIA"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPPO_NR_ID_PLANO_COMUNICACAO"	
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_SINER"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_LEG_TEC"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_LEG_FUN"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_GESTOR_PROC"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_TX_OPERACAO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_CD_NR_USUARIO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_DT_ATUALIZACAO"		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " FROM XPEP_PLANO_TAREFA_PPO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PLTA_NR_SEQUENCIA_TAREFA = " & rds_sqlGeralAlteracao("PLTA_NR_SEQUENCIA_TAREFA")
		
		'Response.write str_sqlAtividadeAlt
		'Response.end
		Set rds_sqlAtividadeAlt = db_Cogest.Execute(str_sqlAtividadeAlt)	
		
		if not rds_sqlAtividadeAlt.eof then		
			str_txtDescrParada 		= Trim(Ucase(rds_sqlAtividadeAlt("PPPO_TX_DESCRICAO_PARADA")))
			str_txtRespSinergia		= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_SINER"))
			str_txtRespLegadoTec	= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_LEG_TEC"))
			str_txtRespLegadoFunc	= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_LEG_FUN"))			 
			str_txtTempParada		= rds_sqlAtividadeAlt("PPPO_TX_QTD_TEMPO_PARADA")
			str_txtUnidTempo 		= Trim(rds_sqlAtividadeAlt("PPPO_TX_UNID_TEMPO_PARADA"))
			str_txtProcedParada 	= Trim(Ucase(rds_sqlAtividadeAlt("PPPO_TX_PROCEDIMENTOS_PARADA")))		
			str_UsuarioGestor		= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_GESTOR_PROC"))
			 			 			 
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtParadaLegado = split(Trim(rds_sqlAtividadeAlt("PPPO_DT_PARADA_LEGADO")),"/")							
			strDia = trim(vetDtParadaLegado(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vetDtParadaLegado(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vetDtParadaLegado(2))
			str_txtDtParadaLegado = strDia & "/" & strMes & "/" & strAno 
			
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtIniR3 = split(Trim(rds_sqlAtividadeAlt("PPPO_DT_INICIO_R3")),"/")							
			strDia = trim(vetDtIniR3(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vetDtIniR3(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vetDtIniR3(2))
			str_txtDtIniR3 = strDia & "/" & strMes & "/" & strAno 
			
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtLimiteAprov = split(Trim(rds_sqlAtividadeAlt("PPPO_DT_LIMITE_APROVACAO")),"/")							
			strDia = trim(vetDtLimiteAprov(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vetDtLimiteAprov(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vetDtLimiteAprov(2))
			str_txtDtLimiteAprov = strDia & "/" & strMes & "/" & strAno

			'str_NomeRespSinergia = RetornaNomeUsuario(str_txtRespSinergia, "Sinergia")	
			'str_NomeRespLegadoTec = RetornaNomeUsuario(str_txtRespLegadoTec, "Legado")
			'str_NomeRespLegadoFunc = RetornaNomeUsuario(str_txtRespLegadoFunc, "Legado")
					
		end if
		rds_sqlAtividadeAlt.close
		set rds_sqlAtividadeAlt = nothing	
	end if
	rds_sqlGeralAlteracao.close
	set rds_sqlGeralAlteracao = nothing		
end if

strReadOnly = ""
if str_Acao ="C" then
	strReadOnly = "readonly"
	str_Texto_Acao = "Consulta"
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
rds_Responsavel.close
set rds_Responsavel = Nothing

'=======================================================================================
'======== ECARREGA DADOS DOS SISTEMAS LEGADOS ==========================================
'str_RespLegado = ""
'str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
'str_RespLegado = str_RespLegado & " USMA_CD_USUARIO "
'str_RespLegado = str_RespLegado & " , USMA_TX_NOME_USUARIO "
'str_RespLegado = str_RespLegado & " FROM dbo.USUARIO_MAPEAMENTO "
'str_RespLegado = str_RespLegado & " Where USMA_TX_MATRICULA <> 0"
'str_RespLegado = str_RespLegado & " ORDER BY USMA_TX_NOME_USUARIO "
'set rds_RespLegado = db_Cogest.Execute(str_RespLegado)
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>:: Cutover - Plano PPO</title>
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
	<script src="../js/global.js" language="javascript"></script>
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
		
	function Verifica_Dif_Numeros(strValor,strNome)	
	{		
		if (isNaN(strValor))
		{
			alert("O contéudo do campo Tempo da Parada deve ser preenchido apenas com nº!");
			
			if (strNome == 'txtTempParada')
			{
				document.frm_Plano_PPO.txtTempParada.value = '';
				document.frm_Plano_PPO.txtTempParada.focus();
				return;
			}
		}
	}	
		
	function confirma_ppo()
	{
		var txt_DescrParada 	= document.frm_Plano_PPO.txtDescrParada.value;			
		var str_RespTecSinGeral = document.frm_Plano_PPO.txtRespTecSinGeral.value; 
		var str_RespTecLegGeral = document.frm_Plano_PPO.txtRespTecLegGeral.value; 
		var str_RespFunLegGeral = document.frm_Plano_PPO.txtRespFunLegGeral.value; 			
		var str_TempParada 		= document.frm_Plano_PPO.txtTempParada.value; 			
		var str_ProcedParada 	= document.frm_Plano_PPO.txtProcedParada.value; 			
		var str_DtParadaLegado	= document.frm_Plano_PPO.txtDtParadaLegado.value;			
		var str_DtIniR3 		= document.frm_Plano_PPO.txtDtIniR3.value;			
		var str_UsuarioGestor 	= document.frm_Plano_PPO.txtUsuarioGestor.value;			
		var str_DtLimiteAprov 	= document.frm_Plano_PPO.txtDtLimiteAprov.value;
		
		//*** Descrição da Parada	
		if (txt_DescrParada == "")
		  {
		  alert("É obrigatório o preenchimento do campo Descrição de Parada!");
		  document.frm_Plano_PPO.txtDescrParada.focus();
		  return;
		  }
		
		//*** Responsável Sinergia				  
		if(str_RespTecSinGeral == '')
		  {
		  alert("É obrigatório o preenchimento do campo Responsável Sinergia!");
		  document.frm_Plano_PPO.txtRespTecSinGeral.focus();
		  return;
		  }
				 
		//*** Responsável Técnico - Legado				  
		if(str_RespTecLegGeral == '')
		  {
		  alert("É obrigatório o preenchimento do campo Responsável Legado - Técnico!");
		  document.frm_Plano_PPO.txtRespTecLegGeral.focus();
		  return;
		  } 
		 
		//*** Responsável Funcional - Legado 
		if(str_RespFunLegGeral == '')
		  {
		  alert("É obrigatório o preenchimento do campo Responsável Legado - Funcional!");
		  document.frm_Plano_PPO.txtRespFunLegGeral.focus();
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
		  alert("É obrigatório o preenchimento do campo Data da Parada do Legado!");
		  document.frm_Plano_PPO.txtDtParadaLegado.focus();
		  return;
		  }
		//else
		 //{
			//validaData(str_DtParadaLegado,'txtDtParadaLegado','Data da Parada do Legado');
			//if (blnData) return;
		 //}
			
	  //*** Data - Início no R/3  	
	   if (str_DtIniR3 == "")
		  {
		  alert("É obrigatório o preenchimento do campo Data - Início no R/3!");
		  document.frm_Plano_PPO.txtDtIniR3.focus();
		  return;
		  }
	   //else
		//{
			//validaData(str_DtIniR3,'txtDtIniR3','Início no R/3');		
			//if (blnData) return;	
		//} 
	
	   //*** Usuário Gestor	
	   if (str_UsuarioGestor == '')
		  {
		  alert("É obrigatório o preenchimento do campo Gestor para o processo!");
		  document.frm_Plano_PPO.txtUsuarioGestor.focus();
		  return;
		  } 
		  
	   //*** Data Limite para Aprovação	
	   if (str_DtLimiteAprov == "")
		  {
		  alert("É obrigatório o preenchimento do campo Data Limite de Aprovação!");
		  document.frm_Plano_PPO.txtDtLimiteAprov.focus();
		  return;
		  }
	   //else
		 //{
			//validaData(str_DtLimiteAprov,'txtDtLimiteAprov','Data Limite para Aprovação');
			//if (blnData) return;	
		 //}  			
	
	   document.frm_Plano_PPO.action="grava_plano.asp?pPlano=PPO";           
	   document.frm_Plano_PPO.submit();				
	}	
	
	function Localiza_Usuario(strTipoResponsavel,strCampo)
	{
		if (strCampo == 'txtRespTecSinGeral')
		{
			strUsuario = document.frm_Plano_PPO.txtRespTecSinGeral.value;		
		
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Responsável Sinergia!");
				document.frm_Plano_PPO.txtRespTecSinGeral.focus();
				return;
			}
		}
		
		if (strCampo == 'txtRespTecLegGeral')
		{
			strUsuario = document.frm_Plano_PPO.txtRespTecLegGeral.value;
			
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Responsável Legado - Técnico!");
				document.frm_Plano_PPO.txtRespTecLegGeral.focus();
				return;
			}		
		}
		
		if (strCampo == 'txtRespFunLegGeral')
		{
			strUsuario = document.frm_Plano_PPO.txtRespFunLegGeral.value;
				
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Responsável Legado - Funcional!");
				document.frm_Plano_PPO.txtRespFunLegGeral.focus();
				return;
			}	
		}
		
		if (strCampo == 'txtUsuarioGestor')
		{
			strUsuario = document.frm_Plano_PPO.txtUsuarioGestor.value;
				
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Gestor do Processo!");
				document.frm_Plano_PPO.txtUsuarioGestor.focus();
				return;
			}	
		}		
		document.frm_Plano_PPO.pTipoResponsavel.value = strTipoResponsavel;	
		document.frm_Plano_PPO.pChaveUsua.value = strUsuario.toUpperCase();	
						
		document.frm_Plano_PPO.action='inclui_altera_plano_ppo.asp?pTipoResponsavel=' + strTipoResponsavel + '&pCampo=' + strCampo;
		document.frm_Plano_PPO.submit();					
	}
	
	function confirma_Exclusao()
	{
		  if(confirm("Confirma a exclusão deste Registro?"))
		  {
			document.frm_Plano_PPO.pAcao.value = 'E';			
			document.frm_Plano_PPO.action='grava_plano.asp?pPlano=PPO' 			        
			document.frm_Plano_PPO.submit();
		  }
	}		
	
	function pega_tamanho(strCampo)
	{	
		if (strCampo == 'txtDescrParada')
		{
			valor = document.forms[0].txtDescrParada.value.length;
			document.forms[0].txttamanhoDescrParada.value = valor;
			if (valor > 100)
			{
				str1 = document.forms[0].txtDescrParada.value;
				str2 = str1.slice(0,100);
				document.forms[0].txtDescrParada.value = str2;
				valor = str2.length;
				document.forms[0].txttamanhoDescrParada.value = valor;
			}
		}
		
		if (strCampo == 'txtProcedParada')
		{
			valor = document.forms[0].txtProcedParada.value.length;
			document.forms[0].txttamanhoProcedParada.value = valor;
			if (valor > 500)
			{
				str1 = document.forms[0].txtProcedParada.value;
				str2 = str1.slice(0,500);
				document.forms[0].txtProcedParada.value = str2;
				valor = str2.length;
				document.forms[0].txttamanhoProcedParada.value = valor;
			}
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
<table width="88%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="6%">&nbsp;</td>
    <td width="80%">&nbsp;</td>
    <td width="14%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo"><table width="100%" border="0" cellpadding="0" cellspacing="7">
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
    <td><table width="95%"  border="0">
      <tr>
        <td><div align="center" class="campo">A&ccedil;&atilde;o</div></td>
      </tr>
      <tr>
        <td bgcolor="#EFEFEF"><div align="center"><span class="campob"><%=str_Texto_Acao%></span></div></td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="75%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr> 
    <td width="19%" bgcolor="#EEEEEE"> <div align="right" class="campo">Atividade:</div></td>
    <td colspan="3" class="campob"><%=str_NomeAtividade%>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Respons&aacute;vel:</div></td>
    <td colspan="3" class="campob"><%=str_Nome_Responsavel%>&nbsp;</td>
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
    <td width="23%" class="campob"><%=dat_Dt_Inicio%>&nbsp;</td>
    <td width="17%" bgcolor="#EEEEEE"> <div align="right" class="campo">Data de 
        T&eacute;rmino:</div></td>
    <td width="41%" class="campob"><%=dat_Dt_Termino%>&nbsp;</td>
  </tr>
</table>
<hr>
<form name="frm_Plano_PPO" method="post">
  <table width="98%" border="0">
    <tr> 
      <td colspan="5"></td>
    </tr>
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="14%" valign="top"><div align="right" class="campob">Descrição da Parada:</div></td>
      <td width="45%">
	  	<%if Request("txtDescrParada") <> "" then%>
	  		<textarea name="txtDescrParada" cols="34" rows="4" <%=strReadOnly%> onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=Request("txtDescrParada")%></textarea></td>
      	<%else%>
			<textarea name="txtDescrParada" cols="34" rows="4" <%=strReadOnly%> onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=str_txtDescrParada%></textarea></td>
		<%end if%>
	  <td width="12%">&nbsp;</td>
      <td width="27%">&nbsp;</td>
    </tr>
	
	  <tr> 
		<td>&nbsp;</td>	
		<td>&nbsp;</td>	
		<td>
			<font face="Verdana" size="2"><font size="1">Caracteres digitados&nbsp;</font><b> 
		  	<input type="text" name="txttamanhoDescrParada" size="5" value="0" maxlength="50" readonly>
		  	</b></font><font face="Verdana" size="1">(Máximo 100 caracteres)</font> 
		</td>
		<td>&nbsp;</td>
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
		if strCampo = "txtRespTecSinGeral" then
			strUsuarioSinergia 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 	
		elseif strCampo = "txtRespTecLegGeral"  then
			strRespTecLegado 	= " - " &RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 	
		elseif strCampo = "txtRespFunLegGeral" then
			strRespFunLegado 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 	
		elseif strCampo = "txtUsuarioGestor" then
			strUsuarioGestor 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel)
		end if				
	end if
	%>	
	
	<%if strUsuarioSinergia <> "" then%>	
		<input type="hidden" value="<%=strUsuarioSinergia%>" name="hdUsuarioSinergia">
	<%else%>
		<input type="hidden" value="<%=Request("hdUsuarioSinergia")%>" name="hdUsuarioSinergia">
	<%end if%>
	<%if strRespTecLegado <> "" then%>	
		<input type="hidden" value="<%=strRespTecLegado%>" name="hdRespTecLegado">
	<%else%>
		<input type="hidden" value="<%=Request("hdRespTecLegado")%>" name="hdRespTecLegado">
	<%end if%>
	<%if strRespFunLegado <> "" then%>	
		<input type="hidden" value="<%=strRespFunLegado%>" name="hdRespFunLegado">
	<%else%>
		<input type="hidden" value="<%=Request("hdRespFunLegado")%>" name="hdRespFunLegado">
	<%end if%>	
	<%if strUsuarioGestor <> "" then%>	
		<input type="hidden" value="<%=strUsuarioGestor%>" name="hdUsuarioGestor">	
	<%else%>
		<input type="hidden" value="<%= Request("hdUsuarioGestor")%>" name="hdUsuarioGestor">	
	<%end if%>		
    <tr>
      <td>&nbsp;</td>
      <td valign="top"><div align="right"><span class="campob">Respons&aacute;vel Sinergia:</span></div></td>
      <td class="campob" valign="bottom">
	    <%if Request("txtRespTecSinGeral") <> "" then %>	  
	  		<input type="text" maxlength="4" value="<%=Request("txtRespTecSinGeral")%>" <%=strReadOnly%> name="txtRespTecSinGeral" size="5">
		<%else%>
			<input type="text" maxlength="4" value="<%=str_txtRespSinergia%>" <%=strReadOnly%> name="txtRespTecSinGeral" size="5">
		<%end if%>
	  	<a href="javascript:Localiza_Usuario('Sinergia','txtRespTecSinGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>
	  	<%if strUsuarioSinergia <> "" then
			Response.write strUsuarioSinergia
		  else
			Response.write Request("hdUsuarioSinergia") 
		end if%>
	  </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>	
    </tr>
    <td height="7" colspan="5"></td>
    </tr>
    <tr> 
      <td colspan="5">
	  	<%'<!--#include file="../includes/inc_lista_Responsavel_Legado.asp" -->%>		
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
				<%if Request("txtRespTecLegGeral") <> "" then%>
					<input type="text" maxlength="4" value="<%=Request("txtRespTecLegGeral")%>" <%=strReadOnly%> name="txtRespTecLegGeral" size="5">
                    <%else%>				
              <input type="text" maxlength="4" value="<%=str_txtRespLegadoTec%>" <%=strReadOnly%> name="txtRespTecLegGeral" size="5">
				<%end if%>
				<a href="javascript:Localiza_Usuario('Legado','txtRespTecLegGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
							
				<%
				if strRespTecLegado <> "" then
					Response.write strRespTecLegado
				else
					Response.write Request("hdRespTecLegado") 
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
					<input type="text" maxlength="4" value="<%=Request("txtRespFunLegGeral")%>" <%=strReadOnly%> name="txtRespFunLegGeral" size="5">
				<%else%>
					<input type="text" maxlength="4" value="<%=str_txtRespLegadoFunc%>" <%=strReadOnly%> name="txtRespFunLegGeral" size="5">
				<%end if%>
				<a href="javascript:Localiza_Usuario('Legado','txtRespFunLegGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
				<%
				if strRespFunLegado <> "" then
					Response.write strRespFunLegado
				else
					Response.write Request("hdRespFunLegado") 
				end if
				%>
			</td>
		  </tr>
		</table>		
	  </td> 
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right" class="campob">Tempo da Parada:</div></td>
      <td class="campo">
	  	<%if Request("txtTempParada") <> "" then%>
			<input name="txtTempParada" type="text" class="txtCampo" size="3" <%=strReadOnly%> onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);" value="<%=Request("txtTempParada")%>">
		<%else%>
			<input name="txtTempParada" type="text" class="txtCampo" size="3" <%=strReadOnly%> onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);" value="<%=str_txtTempParada%>">
		<%end if%>
			
		<%if str_Acao <> "C" then%>	    
			<select name="selUnidMedida" size="1" class="cmd150">          
			  <%if str_txtUnidTempo = "Hora" then%>
				<option value="Hora" selected>Hora</option>
			  <%elseif Request("selUnidMedida") = "Hora" then%>
				<option value="Hora" selected>Hora</option>
			  <%else%>
				<option value="Hora">Hora</option>
			  <%end if%>
			  
			  <%if str_txtUnidTempo = "Dia Útil" then%>
				<option value="Dia Útil" selected>Dia Útil</option>
			  <%elseif Request("selUnidMedida") = "Dia Útil" then%>
				<option value="Dia Útil" selected>Dia Útil</option>
			  <%else%>
				<option value="Dia Útil">Dia Útil</option>
			  <%end if%>
			  
			  <%if str_txtUnidTempo = "Dia Corrido" then%>
				<option value="Dia Corrido" selected>Dia Corrido</option>
			  <%elseif Request("selUnidMedida") = "Dia Corrido" then%>
				<option value="Dia Corrido" selected>Dia Corrido</option>
			  <%else%>
				<option value="Dia Corrido">Dia Corrido</option>
			  <%end if%>
			  
			   <%if str_txtUnidTempo = "Mês" then%>
				<option value="Mês" selected>Mês</option>
			   <%elseif Request("selUnidMedida") = "Mês" then%>
				<option value="Mês" selected>Mês</option>
			   <%else%>
				<option value="Mês">Mês</option>
			   <%end if%>         
			</select>
		<%else
			'*** Mostra na Tela o Tempo de Parada 
			if str_txtUnidTempo <> "" then
				Response.write str_txtUnidTempo
			else
				Response.write Request("selUnidMedida") 
			end if			
		end if%>
	  </td>
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
      <td>
	  	 <%if Request("txtProcedParada") <> "" then%>
	  		<textarea name="txtProcedParada" cols="34" rows="4" <%=strReadOnly%> onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=Request("txtProcedParada")%></textarea></td>
      	 <%else%>
		 	<textarea name="txtProcedParada" cols="34" rows="4" <%=strReadOnly%> onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=str_txtProcedParada%></textarea></td>
		 <%end if%>
	  <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	
	<tr> 
		<td>&nbsp;</td>	
		<td>&nbsp;</td>	
		<td>
			<font face="Verdana" size="2"><font size="1">Caracteres digitados&nbsp;</font><b> 
		  	<input type="text" name="txttamanhoProcedParada" size="5" value="0" maxlength="50" readonly>
		  	</b></font><font face="Verdana" size="1">(Máximo 500 caracteres)</font> 
		</td>
		<td>&nbsp;</td>
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
      <td><div align="right" class="campob">Data da Parada do Legado:</div></td>
      <td>        <table width="100%"  border="0">
          <tr>
            <td>            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="47%">
					<%if Request("txtDtParadaLegado") <> "" then%>
						<input name="txtDtParadaLegado" type="text" class="txtCampo" size="10" maxlength="10" value="<%=Request("txtDtParadaLegado")%>" readonly>
					<%else%>
						<input name="txtDtParadaLegado" type="text" class="txtCampo" size="10" maxlength="10" value="<%=str_txtDtParadaLegado%>" readonly>
					<%end if%>
				</td>
                <td width="53%">
					<%if str_Acao <> "C" then%> 	
						<a href="javascript:show_calendar(true,'frm_Plano_PPO.txtDtParadaLegado','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a></td>
              		<%end if%>
			  </tr>
            </table></td>
            <td><div align="right"><span class="campob">Início no R/3:</span></div></td>
            <td>              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="46%">
				  	<%if Request("txtDtIniR3") <> "" then%>
						<input name="txtDtIniR3" type="text" class="txtCampo" size="10" maxlength="10" value="<%=Request("txtDtIniR3")%>" readonly>
					<%else%>
						<input name="txtDtIniR3" type="text" class="txtCampo" size="10" maxlength="10" value="<%=str_txtDtIniR3%>" readonly>
					<%end if%>				  	
				  </td>
                  <td width="54%">
				  	<%if str_Acao <> "C" then%>
				  		<a href="javascript:show_calendar(true,'frm_Plano_PPO.txtDtIniR3','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a></td>
                	<%end if%>
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
          <div align="right">Gestor do Processo:</div>
        </div></td>
      <td class="campob">
	  	<%if Request("txtUsuarioGestor") <> "" then%>
			<input type="text" maxlength="4" value="<%=Request("txtUsuarioGestor")%>" <%=strReadOnly%> name="txtUsuarioGestor" size="5">
		<%else%>	  	
			<input type="text" maxlength="4" value="<%=str_UsuarioGestor%>" <%=strReadOnly%> name="txtUsuarioGestor" size="5">	
		<%end if%>
		<a href="javascript:Localiza_Usuario('Legado','txtUsuarioGestor');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
		<%
		if strUsuarioGestor <> "" then
			Response.write strUsuarioGestor
		else
			Response.write Request("hdUsuarioGestor") 
		end if
		%>
	  </td>
      <td align="right" valign="top" class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td align="right" valign="top" class="campob"><div align="right">Data Limite para Aprova&ccedil;&atilde;o:</div></td>
      <td>
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="15%">
				<%if Request("txtDtLimiteAprov") <> "" then%>
					<input name="txtDtLimiteAprov" type="text" class="txtCampo" size="10" maxlength="10" onFocus="document.sample.button2.focus()" value="<%=Request("txtDtLimiteAprov")%>" readonly>
				<%else%>
					<input name="txtDtLimiteAprov" type="text" class="txtCampo" size="10" maxlength="10" onFocus="document.sample.button2.focus()" value="<%=str_txtDtLimiteAprov%>" readonly>
				<%end if%>				
			</td>
            <td width="85%">
				<%if str_Acao <> "C" then%>
					<a href="javascript:show_calendar(true,'frm_Plano_PPO.txtDtLimiteAprov','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a></td>
         		<%end if%>
		  </tr>
      </table></td>
      <td>
	  	<input type="hidden" value="<%=int_Cd_ProjetoProject%>" name="pCdProjProject">
	  	<input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">        
		<input type="hidden" value="<%=int_Plano%>" name="pintPlano">
		<input type="hidden" value="<%=int_IdAtividade%>" name="pIdAtividade">		
        <input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
        <input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
        <input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">
        <input type="hidden" value="<%=str_Acao%>" name="pAcao">
		<input type="hidden" value="<%=int_CD_Onda%>" name="pOnda">
		<input type="hidden" value="" name="pTipoResponsavel">
		<input type="hidden" value="" name="pChaveUsua">		
	  </td>
      <td>&nbsp;</td>
    </tr>
	<% if str_Acao = "A" then%>
   <tr> 
      <td>&nbsp;</td>
	   <td colspan="3" align="left" valign="bottom">
	    <div class="campob">
	      <table width="68%" border="0">
            <tr>
              <td width="93%">Link com Plano de A&ccedil;&otilde;es Corretivas / Conting&ecirc;ncia (PAC):</td>
              <td width="7%"><div class="campob">
			  	<%if str_Acao <> "C" then%>
			  		<a href="encaminha_plano.asp?selTipoCadastro=PAC&pSiglaPlano=PAC&pAtividade_Origen=<%="PPO - " & str_NomeAtividade%>&selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Plano%>|0&selTask1=<%=int_IdAtividade%>|<%=int_ResData%>&selTaskSub=<%=int_Id_TarefaProject%>|0"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></div></td>
           		<%end if%>
		    </tr>
          </table>
        </div>
	  	</td>	      
      <td>&nbsp;</td>
    </tr>
	<% end if %>
  </table>
  <table width="625" border="0" align="center">
    <tr>
      <td>&nbsp;</td>
      <td width="1">&nbsp;</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td>&nbsp;</td>
      <td><div align="center" class="campo"></div></td>
    </tr>
    <tr>
      <td width="85">	  	
		<%if str_Acao <> "C" then%>
	  		<a href="javascript:confirma_ppo()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
		<%end if%>
	  <td width="22"><b></b></td>
      <td width="186">
	    <%if str_Acao = "A" and str_Acao <> "C" then%>
			<a href="javascript:confirma_Exclusao();"><img src="../img/botao_excluir.gif" width="85" height="19" border="0"></a>
		<%end if%>
	  </td>
      <td width="165"><%if str_Acao = "C" then%>
        <a href="javascript:history.go(-1)"><img src="../img/botao_voltar.gif" width="85" height="19" border="0"></a>
        <%end if%></td>
      <td width="10">&nbsp;</td>
      <td width="9"></td>
      <td width="8">&nbsp;</td>
      <td width="110"><div align="center"></div></td>
    </tr>
  </table>
</form>
<%
db_Cronograma.close
set db_Cronograma = nothing

db_Cogest.close
set db_Cogest = nothing
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