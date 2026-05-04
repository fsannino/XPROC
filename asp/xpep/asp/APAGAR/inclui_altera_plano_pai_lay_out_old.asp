<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Expires=0

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

'session("CD_Onda") 		= int_CD_Onda
'session("CD_Plano1") 	= int_Plano
'session("CD_Plano2") 	= ""
'session("CD_Fase") 		= str_Fase

'Response.write str_Fase
'Response.end

if str_Acao = "I" then
   str_Texto_Acao = "Inclusão"
else
    str_Texto_Acao = "Alteração"   
   
    str_sqlGeralAlteracao = ""
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & " FROM XPEP_PLANO_TAREFA_GERAL"
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & " WHERE PLTA_NR_ID_TAREFA_PROJECT = " & int_Id_TarefaProject	
	Set rds_sqlGeralAlteracao = db_Cogest.Execute(str_sqlGeralAlteracao)	
			
	'Response.write str_sqlGeralAlteracao & "<br><br><br>"
	'Response.end
	if not rds_sqlGeralAlteracao.eof then		
		str_sqlAtividadeAlt = ""
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & "SELECT PLAN_NR_SEQUENCIA_PLANO"			
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PLTA_NR_SEQUENCIA_TAREFA "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_CD_INTERFACE "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_GRUPO "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_TIPO_PROCESSAMENTO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_NOME_INTERFACE "				
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_PROGRAMA_ENVOLVIDO "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_PRE_REQUISITO "	
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_RESTRICAO "	
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_DEPENDENCIA "			
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_RESP_ACIONAMENTO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_DT_INICIO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_TX_PROCEDIMENTO "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_NR_ID_PLANO_CONTINGENCIA "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAI_NR_ID_PLANO_COMUNICACAO "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_SINER "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_LEG "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_TX_OPERACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_CD_NR_USUARIO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_DT_ATUALIZACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " FROM XPEP_PLANO_TAREFA_PAI"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PLTA_NR_SEQUENCIA_TAREFA = " & int_Id_TarefaProject
						
		'Response.write str_sqlAtividadeAlt
		'Response.end										
		Set rds_sqlAtividadeAlt = db_Cogest.Execute(str_sqlAtividadeAlt)	
		
		rds_sqlGeralAlteracao.close
		set rds_sqlGeralAlteracao = nothing
		
		if not rds_sqlAtividadeAlt.eof then					
			
			str_txtCdInterface 		= Trim(rds_sqlAtividadeAlt("PPAI_TX_CD_INTERFACE"))			
			str_txtGrupo			= Trim(rds_sqlAtividadeAlt("PPAI_TX_GRUPO"))	
			str_TipoBatch			= Trim(rds_sqlAtividadeAlt("PPAI_TX_TIPO_PROCESSAMENTO"))		
			str_txtNomeInterface	= Trim(rds_sqlAtividadeAlt("PPAI_TX_NOME_INTERFACE"))	
			str_txtNomeInterface	= Trim(rds_sqlAtividadeAlt("PPAI_TX_NOME_INTERFACE"))	
			str_txtPgrmEnvolv		= Trim(rds_sqlAtividadeAlt("PPAI_TX_PROGRAMA_ENVOLVIDO"))					
			str_txtPreRequisitos	= Trim(rds_sqlAtividadeAlt("PPAI_TX_PRE_REQUISITO"))		
			str_txtRestricoes		= Trim(rds_sqlAtividadeAlt("PPAI_TX_RESTRICAO"))			
			str_txtDependencias		= Trim(rds_sqlAtividadeAlt("PPAI_TX_DEPENDENCIA"))
			str_txtRespAciona		= Trim(rds_sqlAtividadeAlt("PPAI_TX_RESP_ACIONAMENTO"))
			str_txtProcedimento		= Trim(rds_sqlAtividadeAlt("PPAI_TX_PROCEDIMENTO"))
			str_RespTecLegGeral		= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_LEG"))
			str_txtRespSinergia		= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_SINER"))
			
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDtInicio_Pai = split(Trim(rds_sqlAtividadeAlt("PPAI_DT_INICIO")),"/")							
			strDia = trim(vetDtInicio_Pai(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vetDtInicio_Pai(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vetDtInicio_Pai(2))
			str_txtDtInicio_Pai = strDia & "/" & strMes & "/" & strAno 			
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
						
						if (strobjnome=='txtPreRequisitos')
						{
							document.forms[0].txtPreRequisitos.value = strvalor.substr(0,i);
						}
						
						if (strobjnome=='txtRestricoes')
						{
							document.forms[0].txtRestricoes.value = strvalor.substr(0,i);
						}
							
						if (strobjnome=='txtProcedimento') 
						{
							document.forms[0].txtProcedimento.value = strvalor.substr(0,i);
						}	
						
						if (strobjnome=='txtDependencias') 
						{
							document.forms[0].txtDependencias.value = strvalor.substr(0,i);
						}
						
						if (strobjnome=='txtProcedimento') 
						{
							document.forms[0].txtProcedimento.value = strvalor.substr(0,i);
						}						
						break;
					}
				}
			}		
		}
				
		function confirma_pai()
		{				
			var str_CdInterface 	= document.frm_Plano_PAI.txtCdInterface.value; 					
			var str_Grupo 			= document.frm_Plano_PAI.txtGrupo.value; 					
			var str_NomeInterface  	= document.frm_Plano_PAI.txtNomeInterface.value;						
			var str_PgrmEnvolv		= document.frm_Plano_PAI.txtPgrmEnvolv.value;			
			var str_PreRequisitos	= document.frm_Plano_PAI.txtPreRequisitos.value;		
			var str_Restricoes		= document.frm_Plano_PAI.txtRestricoes.value;		
			var str_Dependencias	= document.frm_Plano_PAI.txtDependencias.value;
			var str_DtInicio_Pai	= document.frm_Plano_PAI.txtDtInicio_Pai.value;		
			var str_RespAciona		= document.frm_Plano_PAI.txtRespAciona.value;
			var str_Procedimento	= document.frm_Plano_PAI.txtProcedimento.value;
			var str_RespTecLegGeral = document.frm_Plano_PAI.txtRespTecLegGeral.value;	
			var str_RespTecSinGeral	= document.frm_Plano_PAI.txtRespTecSinGeral.value;	
										  
			//*** Código da Interface
			if (str_CdInterface == "")
			  {
			  alert("É obrigatório o preenchimento do campo Código da Interface!");
			  document.frm_Plano_PAI.txtCdInterface.focus();
			  return;
			  }		
									 
		   //*** Grupo
		   /*if (str_Grupo == "")
			  {
			  alert("É obrigatório o preenchimento do campo Grupo!");
			  document.frm_Plano_PAI.txtGrupo.focus();
			  return;
			  }
							*/
		   //*** Nome da Interface
		   if (str_NomeInterface == "")
			  {
			  alert("É obrigatório o preenchimento do campo Nome da Interface!");
			  document.frm_Plano_PAI.txtNomeInterface.focus();
			  return;
			  } 
				
		  //*** Programa Envolvido
		   if (str_PgrmEnvolv == "")
			  {
			  alert("É obrigatório o preenchimento do campo Programa Envolvido!");
			  document.frm_Plano_PAI.txtPgrmEnvolv.focus();
			  return;
			  } 
		
		   //*** Pré-Requisitos
		   if (str_PreRequisitos == "")
			  {
			  alert("É obrigatório o preenchimento do campo Pré-Requisitos!");
			  document.frm_Plano_PAI.txtPreRequisitos.focus();
			  return;
			  } 		
						
		   //*** Restrições
		   if (str_Restricoes == "")
			  {
			  alert("É obrigatório o preenchimento do campo Restrições!");
			  document.frm_Plano_PAI.txtRestricoes.focus();
			  return;
			  }   
			
		   //*** Dependências
		   if (str_Dependencias == "")
			  {
			  alert("É obrigatório o preenchimento do campo Dependências!");
			  document.frm_Plano_PAI.txtDependencias.focus();
			  return;
			  } 
			  
		   //*** Data de Inicio
		   if (str_DtInicio_Pai == "")
			  {
			  alert("É obrigatório o preenchimento do campo Data de Início!");
			  document.frm_Plano_PAI.txtDtInicio_Pai.focus();
			  return;
			  }  
		  // else
			//  {
				//validaData(str_DtInicio_Pai,'txtDtInicio_Pai','Data de Início');
				//if (blnData) return; 
			 // }  		
			  
		   //*** Responsável pelo Acionamento
		   if (str_RespAciona == "")
			  {
			  alert("É obrigatório o preenchimento do campo Responsável pelo Acionamento!");
			  document.frm_Plano_PAI.txtRespAciona.focus();
			  return;
			  }   
				
		   //*** Procedimento
		   if (str_Procedimento == "")
			  {
			  alert("É obrigatório o preenchimento do campo Procedimento!");
			  document.frm_Plano_PAI.txtProcedimento.focus();
			  return;
			  }    	
					
		   //*** Responsável Legado
		   if (str_RespTecLegGeral == "")
			  {
			  alert("É obrigatório o preenchimento do campo Responsável Legado!");
			  document.frm_Plano_PAI.txtRespTecLegGeral.focus();
			  return;
			  }    
			  
		   //*** Responsável Sinergia
		   if (str_RespTecSinGeral == "")
			  {
			  alert("É obrigatório o preenchimento do campo Responsável Sinergia!");
			  document.frm_Plano_PAI.txtRespTecSinGeral.focus();
			  return;
			  }    	
			  
		   document.frm_Plano_PAI.action="grava_plano.asp?pPlano=PAI";           
		   document.frm_Plano_PAI.submit();				
		}	
		
		function Localiza_Usuario(strTipoResponsavel,strCampo)
		{		
			if (strCampo == 'txtRespAciona')
			{
				strUsuario = document.frm_Plano_PAI.txtRespAciona.value;		
			
				if (strUsuario == '')
				{			
					alert("É obrigatório o preenchimento do campo Responsável pelo Acionamento!");
					document.frm_Plano_PAI.txtRespAciona.focus();
					return;
				}
			}
			
			if (strCampo == 'txtRespTecLegGeral')
			{
				strUsuario = document.frm_Plano_PAI.txtRespTecLegGeral.value;
				
				if (strUsuario == '')
				{			
					alert("É obrigatório o preenchimento do campo Responsável Legado!");
					document.frm_Plano_PAI.txtRespTecLegGeral.focus();
					return;
				}		
			}			
						
			if (strCampo == 'txtRespTecSinGeral')
			{
				strUsuario = document.frm_Plano_PAI.txtRespTecSinGeral.value;
				
				if (strUsuario == '')
				{			
					alert("É obrigatório o preenchimento do campo Responsável Sinergia!");
					document.frm_Plano_PAI.txtRespTecSinGeral.focus();
					return;
				}		
			}					
						
						
						
			document.frm_Plano_PAI.pTipoResponsavel.value = strTipoResponsavel;	
			document.frm_Plano_PAI.pChaveUsua.value = strUsuario.toUpperCase();	
							
			document.frm_Plano_PAI.action='inclui_altera_plano_pai.asp?pTipoResponsavel=' + strTipoResponsavel + '&pCampo=' + strCampo;
			document.frm_Plano_PAI.submit();		
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
    <td width="81%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo"><table width="100%" border="0" cellpadding="0" cellspacing="7">
        <tr> 
          <td width="8%"><div align="right" class="subtitulob">Onda:</div></td>
          <td colspan="2" class="subtitulo"><%=str_Desc_Onda%></td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td width="8%" class="subtitulob">Plano:</td>
          <td width="84%" class="subtitulo">Plano de Acionamento de Interfaces e Processos BATCH - PAI</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="75%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr> 
    <td width="17%" bgcolor="#EEEEEE"><div align="right" class="campo">Atvidade:</div></td>
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
    <td width="21%" class="campob"><%=dat_Dt_Inicio%>&nbsp;</td>
    <td width="20%" bgcolor="#EEEEEE"><div align="right" class="campo">Data de T&eacute;rmino:</div></td>
    <td width="33%" class="campob"><%=dat_Dt_Termino%>&nbsp;</td>
  </tr>
</table>
<form name="frm_Plano_PAI" method="post" action="">
  <table width="98%" border="0">      
	<tr> 
	  <td colspan="5"><hr></td>
	</tr>      
	    
    <tr>
		<td width="1%" class="campo">&nbsp;</td>
		<td width="19%" class="campob"><div align="right">Código da Interface:</div></td>
		<td width="39%"><input type="text" size="45" maxlength="50" name="txtCdInterface" value="<%=str_txtCdInterface%>"></td>
		<td width="4%">&nbsp;</td>
		<td width="37%">&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Grupo:</div></td>
		<td><input type="text" size="45" maxlength="50" name="txtGrupo" value="<%=str_txtGrupo%>"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    <tr> 
		<td class="campo">&nbsp;</td>
		<td height="25" class="campob"><div align="right">Tipo:</div></td>
		<td class="campo"> 
			<select name="selTipoBatch">
				<%if str_TipoBatch = "Batch" then%>
			  		<option value="Batch" selected>Batch</option>
				<%else%>
					<option value="Batch">Batch</option>
				<%end if%>
				<%if str_TipoBatch = "On-Line" then%>
			  		<option value="On-Line" selected>On-Line</option>
				<%else%>
					<option value="On-Line">On-Line</option>
				<%end if%>			  
			</select> 
	  </td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Nome da Interface:</div></td>
		<td><input type="text" size="45" maxlength="70" name="txtNomeInterface" value="<%=str_txtNomeInterface%>"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Programa Envolvido:</div></td>
		<td><input type="text" size="45" maxlength="50" name="txtPgrmEnvolv" value="<%=str_txtPgrmEnvolv%>"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob" valign="top"><div align="right">Pré-Requisitos:</div></td>
		<td><textarea type="text" cols="34" rows="4" name="txtPreRequisitos" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtPreRequisitos%></textarea></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob" valign="top"><div align="right">Restrições:</div></td>
		<td><textarea type="text" cols="34" rows="4" name="txtRestricoes" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtRestricoes%></textarea></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
        
    </tr>
		<td class="campo">&nbsp;</td>
		<td valign="top" class="campob"><div align="right">Dependências:</div></td>
		<td>
			<!--<input type="text" name="txtDependencias" value="<%'=str_txtDependencias%>">-->
			<textarea type="text" cols="34" rows="4" name="txtDependencias" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtDependencias%></textarea>
		</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Data de Inicio:</div></td>
		<td>
			<input type="text" name="txtDtInicio_Pai" maxlength="10" size="10" value="<%=str_txtDtInicio_Pai%>">
			<script>
				var a = document.height
			</script>
			<a href="javascript:show_calendar(true,'frm_Plano_PAI.txtDtInicio_Pai','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" id="img1" width="24" height="22" border="0"></a>
		</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
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
	
	'Response.write "strTipoResponsavel " & strTipoResponsavel
	
	if strTipoResponsavel <> "" then				
		if strCampo = "txtRespAciona" then
			strResAcionamento 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 	
		elseif strCampo = "txtRespTecLegGeral"  then
			strRespTecLeg		= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel)		
		elseif strCampo = "txtRespTecSinGeral"  then
			strRespSinergia		= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel)
		end if				
	end if%>	
	
	<%if strResAcionamento <> "" then%>	
		<input type="hidden" value="<%=strResAcionamento%>" name="hdResAcionamento">
	<%else%>
		<input type="hidden" value="<%=Request("hdResAcionamento")%>" name="hdResAcionamento">
	<%end if%>
	
	<%if strRespTecLeg <> "" then%>	
		<input type="text" value="<%=strRespTecLeg%>" name="hdRespTecLegGeral">
	<%else%>
		<input type="text" value="<%=Request("hdRespTecLegGeral")%>" name="hdRespTecLegGeral">
	<%end if%>	
	
	<%if strRespSinergia <> "" then%>	
		<input type="hidden" value="<%=strRespSinergia%>" name="hdRespSinergia">
	<%else%>
		<input type="hidden" value="<%=Request("hdRespSinergia")%>" name="hdRespSinergia">
	<%end if%>		
	
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Responsável pelo Acionamento:</div></td>
		<td class="campob">
			<%if Request("txtRespAciona") <> "" then%>	  
				<input type="text" maxlength="4" value="<%=Request("txtRespAciona")%>" name="txtRespAciona" size="5">				
			<%else%>
				<input type="text" maxlength="4" value="<%=str_txtRespAciona%>" name="txtRespAciona" size="5">
			<%end if%>
			<a href="javascript:Localiza_Usuario('Legado','txtRespAciona');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>
			<%if strResAcionamento <> "" then
				Response.write strResAcionamento
			  else
				Response.write Request("hdResAcionamento") 
			  end if%>			
		</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
        </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob" valign="top"><div align="right">Procedimento:</div></td>
		<td><textarea type="text" cols="34" rows="4" name="txtProcedimento" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtProcedimento%></textarea></td>
		<td>
		  <input type="hidden" value="<%=int_Cd_ProjetoProject%>" name="pCdProjProject">
		  <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
          <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
          <input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
          <input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
          <input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">          
		  <input type="hidden" value="<%=str_Acao%>" name="pAcao">
		  <input type="hidden" value="<%=int_CD_Onda%>" name="pOnda">
		  <input type="hidden" value="<%=str_Fase%>" name="pFase">
		  <input type="hidden" value="<%=strPlanoOriginal%>" name="pPlanoOriginal">
		  <input type="hidden" value="" name="pTipoResponsavel">
		  <input type="hidden" value="" name="pChaveUsua">	
		  </td>	
		<td>&nbsp;</td>
    </tr>
    
    <tr> 
      <td colspan="5">    <tr> 
      <td colspan="5"><table width="100%"  border="0">
  <tr>
    <td>&nbsp;</td>
    <td class="campob" align="right">Respons&aacute;vel Legado:</td>
    <td class="campob">		
		<%if Request("txtRespTecLegGeral") <> "" then%>	  
			<input type="text" maxlength="4" value="<%=Request("txtRespTecLegGeral")%>" name="txtRespTecLegGeral" size="5"> 
		<%else%>
			<input type="text" maxlength="4" value="<%=str_RespTecLegGeral%>" name="txtRespTecLegGeral" size="5"> 
		<%end if%>
		<a href="javascript:Localiza_Usuario('Legado','txtRespTecLegGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>
		<%if strRespTecLeg <> "" then
			Response.write strRespTecLeg
		  else
			Response.write Request("hdRespTecLegGeral") 
		end if%>			
	</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="1%">&nbsp;</td>
    <td width="18%" class="campob" align="right">Respons&aacute;vel Sinergia:</td>
    <td width="79%" class="campob">		
		<%if Request("txtRespTecSinGeral") <> "" then%>	  
			<input type="text" maxlength="4" value="<%=str_txtRespSinergia%>" name="txtRespTecSinGeral" size="5">
		<%else%>
			<input type="text" maxlength="4" value="<%=str_txtRespSinergia%>" name="txtRespTecSinGeral" size="5">
		<%end if%>
		<a href="javascript:Localiza_Usuario('Sinergia','txtRespTecSinGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>
		<%if strRespSinergia <> "" then
			Response.write strRespSinergia
		  else
			Response.write Request("hdRespSinergia") 
		end if%>			
	</td>
    <td width="1%">&nbsp;</td>
    <td width="1%">&nbsp;</td>
  </tr>
</table>
      </td>
    </tr>   
	<%if str_Acao = "A" then%> 
    <tr>
      <td>&nbsp;</td>
      <td><div align="right" class="campob">Procedimentos de Conting&ecirc;ncia:</div></td>
      <td><a href="encaminha_plano.asp?selTipoCadastro=PAC&pSiglaPlano=PAC&pAtividade_Origen=<%="PAI - " & str_NomeAtividade%>&selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Plano%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
      <td class="campob">&nbsp;</td>
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
      <td width="85"><a href="javascript:confirma_pai()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
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