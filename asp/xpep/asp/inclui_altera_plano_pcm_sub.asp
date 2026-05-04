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
intResData = request("pResData")

if Request("pintPlano") <> "" then
	int_Plano = Request("pintPlano")
else
	int_Plano = request("pPlano")
end if

if Request("pintPlano2") <> "" then
	int_Plano2 = request("pintPlano2")
else
	int_Plano2 = request("pPlano2")
end if

strNomePlanoOrigem = request("pPlano_Origem")
str_CdSeqPCM = request("pCdSeqPCM")

'Response.write "str_Acao " & str_Acao & "<br>"
'Response.write "int_Cd_ProjetoProject " & int_Cd_ProjetoProject & "<br>"
'Response.write "int_Id_TarefaProject " & int_Id_TarefaProject & "<br>"
'Response.write "int_CD_Onda " & int_CD_Onda & "<br>"
'Response.write "int_Plano " & int_Plano & "<br>"
'Response.write "int_Plano2 " & int_Plano2 & "<br>"
'Response.write "strNomePlanoOrigem " & strNomePlanoOrigem & "<br>"
'Response.write "str_CdSeqPCM " & str_CdSeqPCM & "<br>"
'Response.end


boo_Escrita = ""
if str_Acao = "I" then
   str_Texto_Acao = "Inclusão"
else
	if str_Acao = "A" then
   		str_Texto_Acao = "Alteração"
   	else
   		str_Texto_Acao = "Consulta"	
		boo_Escrita = "readonly"
	end if
   intCdSeqPCM = Request("pCdSeqPCM")
   
    'str_sqlGeralAlteracao = ""
	'str_sqlGeralAlteracao = str_sqlGeralAlteracao & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
	'str_sqlGeralAlteracao = str_sqlGeralAlteracao & " FROM XPEP_PLANO_TAREFA_GERAL"
	'str_sqlGeralAlteracao = str_sqlGeralAlteracao & " WHERE PLTA_NR_ID_TAREFA_PROJECT = " & int_Id_TarefaProject	
	'Set rds_sqlGeralAlteracao = db_Cogest.Execute(str_sqlGeralAlteracao)	
			
	'Response.write str_sqlGeralAlteracao & "<br><br><br>"
	'Response.end
	'if not rds_sqlGeralAlteracao.eof then		
		str_sqlAtividadeAlt = ""
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & "SELECT PLAN_NR_SEQUENCIA_PLANO"		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_NR_SEQUENCIA_TAREFA "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_ATIVIDADE "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_TP_COMUNICACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_O_QUE_COMUNICAR "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_PARA_QUEM "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_UNID_ORGAO "				
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_QUANDO_OCORRE "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_RESP_CONTEUDO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_RESP_DIVULGACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_COMO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_APROVADOR_PB "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_DT_APROVACAO "	
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_ARQUIVO_ANEXO1 "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_ARQUIVO_ANEXO2 "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPCM_TX_ARQUIVO_ANEXO3 "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_TX_OPERACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_CD_NR_USUARIO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_DT_ATUALIZACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " FROM XPEP_PLANO_TAREFA_PCM"
		'str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PLTA_NR_SEQUENCIA_TAREFA = " & rds_sqlGeralAlteracao("PLTA_NR_SEQUENCIA_TAREFA")
		'str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PLTA_NR_SEQUENCIA_TAREFA = " & int_Id_TarefaProject
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PPCM_NR_SEQUENCIA_TAREFA = " & intCdSeqPCM
						
		'Response.write str_sqlAtividadeAlt
		'Response.end										
		Set rds_sqlAtividadeAlt = db_Cogest.Execute(str_sqlAtividadeAlt)	
		
		'rds_sqlGeralAlteracao.close
		'set rds_sqlGeralAlteracao = nothing
		
		if not rds_sqlAtividadeAlt.eof then		

			str_Atividade 		= Trim(rds_sqlAtividadeAlt("PPCM_TX_ATIVIDADE"))	
			str_Comunicacao 	= Trim(rds_sqlAtividadeAlt("PPCM_TX_TP_COMUNICACAO"))
			str_OqueComunicar 	= Trim(rds_sqlAtividadeAlt("PPCM_TX_O_QUE_COMUNICAR"))
			str_AQuemComunicar	= Trim(rds_sqlAtividadeAlt("PPCM_TX_PARA_QUEM"))	
			str_UnidadeOrgao	= Trim(rds_sqlAtividadeAlt("PPCM_TX_UNID_ORGAO"))		
			str_RespConteudo 	= Trim(rds_sqlAtividadeAlt("PPCM_TX_RESP_CONTEUDO"))
			str_RespDivulg		= Trim(rds_sqlAtividadeAlt("PPCM_TX_RESP_DIVULGACAO"))
			str_Como			= Trim(rds_sqlAtividadeAlt("PPCM_TX_COMO"))
			
			str_arquivo1		= Trim(rds_sqlAtividadeAlt("PPCM_TX_ARQUIVO_ANEXO1"))
			str_arquivo2		= Trim(rds_sqlAtividadeAlt("PPCM_TX_ARQUIVO_ANEXO2"))
			str_arquivo3		= Trim(rds_sqlAtividadeAlt("PPCM_TX_ARQUIVO_ANEXO3"))
			
			str_AprovadorPB		= Trim(rds_sqlAtividadeAlt("PPCM_TX_APROVADOR_PB"))							
			

			strDia = ""		
			strMes = ""
			strAno = ""
			vet_QuandoOcorre = split(Trim(rds_sqlAtividadeAlt("PPCM_TX_QUANDO_OCORRE")),"/")							
			strDia = trim(vet_QuandoOcorre(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vet_QuandoOcorre(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vet_QuandoOcorre(2))
			str_txtQuandoOcorre = strDia & "/" & strMes & "/" & strAno
			
			if not IsNull(rds_sqlAtividadeAlt("PPCM_DT_APROVACAO")) then
				strDia = ""		
				strMes = ""
				strAno = ""
				vetDtAprovacao = split(Trim(rds_sqlAtividadeAlt("PPCM_DT_APROVACAO")),"/")							
				strDia = trim(vetDtAprovacao(1))
				if cint(strDia) < 10 then
					strDia = "0" & strDia
				end if			
				strMes = trim(vetDtAprovacao(0))			
				if cint(strMes) < 10 then
					strMes = "0" & strMes
				end if
				strAno = trim(vetDtAprovacao(2))
				str_txtDtAprovacao = strDia & "/" & strMes & "/" & strAno
			else
   			    str_txtDtAprovacao = ""
			end if	 			
		end if	
		rds_sqlAtividadeAlt.close
		set rds_sqlAtividadeAlt = nothing
	'end if 
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
'str_Sql_DadosAdicionais_Tarefa = ""
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " SELECT   "
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " TASK_UID"
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " , TASK_NAME"
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " , RESERVED_DATA"
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " , TASK_START_DATE"
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " , TASK_FINISH_DATE"
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " FROM MSP_TASKS"
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " WHERE PROJ_ID = " & int_Cd_ProjetoProject
'str_Sql_DadosAdicionais_Tarefa = str_Sql_DadosAdicionais_Tarefa & " AND TASK_UID = " & int_Id_TarefaProject 
'set rds_DadosAdicionais_Tarefa = db_Cronograma.Execute(str_Sql_DadosAdicionais_Tarefa)
'if not rds_DadosAdicionais_Tarefa.Eof then
'   dat_Dt_Inicio = rds_DadosAdicionais_Tarefa("TASK_START_DATE")
'   dat_Dt_Termino = rds_DadosAdicionais_Tarefa("TASK_FINISH_DATE")   
'   str_NomeAtividade = rds_DadosAdicionais_Tarefa("TASK_NAME")
'else
'   dat_Dt_Inicio = ""
'   dat_Dt_Termino = ""
'end if
'rds_DadosAdicionais_Tarefa.close
'set rds_DadosAdicionais_Tarefa = Nothing


'=======================================================================================
' ===== ENCONTRA RESPONSÁVEL PELA TAREFA ===============================================
'str_Responsavel = ""
'str_Responsavel = str_Responsavel & " SELECT MSP_TEXT_FIELDS.TEXT_VALUE "
'str_Responsavel = str_Responsavel & " FROM MSP_TEXT_FIELDS "
'str_Responsavel = str_Responsavel & " INNER JOIN MSP_CONVERSIONS ON MSP_TEXT_FIELDS.TEXT_FIELD_ID = MSP_CONVERSIONS.CONV_VALUE"
'str_Responsavel = str_Responsavel & " WHERE MSP_CONVERSIONS.CONV_STRING='Task Text11'"
'str_Responsavel = str_Responsavel & " AND MSP_TEXT_FIELDS.PROJ_ID = " & int_Cd_ProjetoProject
'str_Responsavel = str_Responsavel & " AND MSP_TEXT_FIELDS.TEXT_REF_UID = " & int_Id_TarefaProject
'set rds_Responsavel = db_Cronograma.Execute(str_Responsavel)
'if not rds_Responsavel.Eof then
'   str_Nome_Responsavel = rds_Responsavel("TEXT_VALUE")
'else
'   str_Nome_Responsavel = " não informado "   
'end if

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
<title>:: Cutover - Plano PCM</title>
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
	<script language="javascript" src="../js/digite-cal.js"></script>		
	<script src="../js/global.js" language="javascript"></script>
	<script language="JavaScript">	
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
					
					if (strobjnome=='txtOqueComunicar')
					{
						document.forms[0].txtOqueComunicar.value = strvalor.substr(0,i);
					}						
					break;
				}
			}
		}		
	}
		
	function confirma_pcm()
	{			
		var str_Atividade 		= document.frm_Plano_PCM.txtAtividade.value; 
		var str_OqueComunicar 	= document.frm_Plano_PCM.txtOqueComunicar.value; 					
		var str_AQuemComunicar 	= document.frm_Plano_PCM.txtAQuemComunicar.value; 
		var str_UnidadeOrgao	= document.frm_Plano_PCM.txtUnidadeOrgao.value;		
		var str_QuandoOcorre  	= document.frm_Plano_PCM.txtQuandoOcorre.value;
						
		var str_RespConteudo	= document.frm_Plano_PCM.txtRespConteudo.value;			
		var str_RespDivulg 		= document.frm_Plano_PCM.txtRespDivulg.value;	
		var str_Como 			= document.frm_Plano_PCM.txtComo.value;						
		var str_AprovadorPB		= document.frm_Plano_PCM.txtAprovadorPB.value;
		var str_txtDtAprovacao	= document.frm_Plano_PCM.txtDtAprovacao.value;						
						
		var str_filArquivo1		= document.frm_Plano_PCM.filArquivo1.value;
		var str_filArquivo2		= document.frm_Plano_PCM.filArquivo2.value;
		var str_filArquivo3		= document.frm_Plano_PCM.filArquivo3.value;
												
		//*** Atividade
		if (str_Atividade == "")
		  {
		  alert("É obrigatório o preenchimento do campo Atividade!");
		  document.frm_Plano_PCM.txtAtividade.focus();
		  return;
		  }	
		  	  
		//*** O que Comunicar
		if (str_OqueComunicar == "")
		  {
		  alert("É obrigatório o preenchimento do campo O que Comunicar!");
		  document.frm_Plano_PCM.txtOqueComunicar.focus();
		  return;
		  }		
						 
	  //*** Para quem comunicar	
	   if (str_AQuemComunicar == "")
		  {
		  alert("É obrigatório o preenchimento do campo Para Quem Comunicar!");
		  document.frm_Plano_PCM.txtAQuemComunicar.focus();
		  return;
		  }
		  
	   //*** Procedimentos da Parada   	
	   if (str_UnidadeOrgao == "")
		  {
		  alert("É obrigatório o preenchimento do campo Unidade de Orgão!");
		  document.frm_Plano_PCM.txtUnidadeOrgao.focus();
		  return;
		  }  								
		
	   //*** Data Limite para Comunicar
	   if (str_QuandoOcorre == "")
		  {
		  alert("É obrigatório o preenchimento do Data Limite para Comunicar!");
		  document.frm_Plano_PCM.txtQuandoOcorre.focus();
		  return;
		  } 
	   //else
		  //{
			//validaData(str_QuandoOcorre,'txtQuandoOcorre','Data Limite para Comunicar');
			//if (blnData) return; 
		 // }  		
	
	/*
	  	if (str_txtDtAprovacao != "")
		  {
			validaData(str_txtDtAprovacao,'txtDtAprovacao','Data de Aprovação');
			if (blnData) return; 
		  }  		
	*/  
		  // **** VERIFICA SE O ARQUIVO 1 FOI SELECIONADO NO DIRETÓRIO "P:"
		  if (str_filArquivo1 != '')
		  {		  	
		  	if (str_filArquivo1.substring(0,2)!='P:')
			{
				alert('O arquivo deve estar localizado no Diretório "P:"');				
				document.frm_Plano_PCM.filArquivo1.focus();
				return;
			}						
		  } 		
		  
		   // **** VERIFICA SE O ARQUIVO 2 FOI SELECIONADO NO DIRETÓRIO "P:"
		  if (str_filArquivo2 != '')
		  {		  	
		  	if (str_filArquivo2.substring(0,2)!='P:')
			{
				alert('O arquivo deve estar localizado no Drive "P"');				
				document.frm_Plano_PCM.filArquivo2.focus();
				return;
			}						
		  } 		
		   // **** VERIFICA SE O ARQUIVO 3 FOI SELECIONADO NO DIRETÓRIO "P:"
		  if (str_filArquivo3 != '')
		  {		  	
		  	if (str_filArquivo3.substring(0,2)!='P:')
			{
				alert('O arquivo deve estar localizado no Diretório "P:"');				
				document.frm_Plano_PCM.filArquivo3.focus();
				return;
			}						
		  }		
		
	   //alert(str_OqueComunicar);	
		
	   document.frm_Plano_PCM.action="grava_plano.asp";           
	   document.frm_Plano_PCM.submit();
	}	
	
	function Localiza_Usuario(strTipoResponsavel,strCampo)
	{	
		if (strCampo == 'txtAprovadorPB')
		{
			strUsuario = document.frm_Plano_PCM.txtAprovadorPB.value;
			
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Aprovador PB!");
				document.frm_Plano_PCM.txtAprovadorPB.focus();
				return;
			}		
		}
				
		document.frm_Plano_PCM.pTipoResponsavel.value = strTipoResponsavel;	
		document.frm_Plano_PCM.pChaveUsua.value = strUsuario.toUpperCase();	
		document.frm_Plano_PCM.pCampo.value = strCampo;
						
		document.frm_Plano_PCM.action='inclui_altera_plano_pcm_sub.asp?pTipoResponsavel=' + strTipoResponsavel + '&pCampo=' + strCampo;
		document.frm_Plano_PCM.submit();			
	}
	
	function pega_tamanho(strCampo)
		{	
			if (strCampo == 'txtAtividade')
			{
				valor = document.forms[0].txtAtividade.value.length;
				document.forms[0].txttamanhoAtividade.value = valor;
				if (valor > 300)
				{
					str1 = document.forms[0].txtAtividade.value;
					str2 = str1.slice(0,300);
					document.forms[0].txtAtividade.value = str2;
					valor = str2.length;
					document.forms[0].txttamanhoAtividade.value = valor;
				}
			}
			
			if (strCampo == 'txtOqueComunicar')
			{
				valor = document.forms[0].txtOqueComunicar.value.length;
				document.forms[0].txttamanhoOqueComunicar.value = valor;
				if (valor > 5000)
				{
					str1 = document.forms[0].txtOqueComunicar.value;
					str2 = str1.slice(0,5000);
					document.forms[0].txtOqueComunicar.value = str2;
					valor = str2.length;
					document.forms[0].txttamanhoOqueComunicar.value = valor;
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
<form name="frm_Plano_PCM" method="post">
  <table width="82%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="1%">&nbsp;</td>
    <td width="85%" class="subtitulo">
	<table width="100%" border="0" cellpadding="0" cellspacing="7">
        <tr>
          <td width="10%"><div align="right"><span class="subtitulob">Onda:</span></div></td>
          <td colspan="3"><%=str_Desc_Onda%></td>
          </tr>
	<% 'if str_Atividade_Origem <> "" then %>
        <tr>
          <td colspan="2"><div align="right"><span class="subtitulob">Plano Origem: </span></div></td>
          <td width="81%" colspan="2" class="subtitulo"><%=strNomePlanoOrigem%></td>
        </tr>
		<% 'end if %>
        <tr> 
          <td colspan="2"><div align="right" class="subtitulob">Plano:</div></td>
          <td colspan="2" class="subtitulo">Plano de Comunicação - PCM</td>
        </tr>
      </table></td>
    <td width="14%"><table width="100%"  border="0">
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center" class="campo">A&ccedil;&atilde;o</div></td>
      </tr>
      <tr>
        <td bgcolor="#EFEFEF"><div align="center"><span class="campob"><%=str_Texto_Acao%></span></div></td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="1">
  <tr>
    <td><hr></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="5" cellpadding="2">
    <tr> 
      <td colspan="3"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Equipe Funcional</strong></font></td>
      <td width="3%">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top" class="campob"><div align="right">Atividade:</div></td>
      <td>
	  	<%if Request("txtAtividade") <> "" then%>
	  		<textarea name="txtAtividade" cols="34" rows="4" id="txtAtividade" <%=boo_Escrita%> onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=Request("txtAtividade")%></textarea>
      	<%else%>
			<textarea name="txtAtividade" cols="34" rows="4" id="txtAtividade" <%=boo_Escrita%> onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=str_Atividade%></textarea>		
		<%end if%>	  	
	  </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
	
	<tr> 
		<td>&nbsp;</td>		
		<td>
			<font face="Verdana" size="2"><font size="1">Caracteres digitados&nbsp;</font><b> 
		  	<input type="text" name="txttamanhoAtividade" size="5" value="0" maxlength="50" readonly>
		  	</b></font><font face="Verdana" size="1">(Máximo 300 caracteres)</font> 
		</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>		
	</tr>
	
    <tr> 
      <td width="19%" class="campob"><div align="right">Comunicação:</div></td>
      <td width="74%" class="campo">
	  <%if str_Acao <> "C" then%>
			<select name="selComunicacao" class="cmdOnda" <%=boo_Escrita%>>
			  <%if str_Comunicacao = "INT" then%>
				<option value="INT" selected>Interna</option>
			  <%elseif Request("selComunicacao") = "INT" then%>
				<option value="INT" selected>Interna</option>
			  <%else%>
				<option value="INT">Interna</option>
			  <%end if%>
			  
			  <%if str_Comunicacao = "EXT" then%>
				<option value="EXT" selected>Externa</option>
			  <%elseif Request("selComunicacao") = "EXT" then%>
				<option value="EXT" selected>EXT</option>
			  <%else%>
				<option value="EXT">Externa</option>
			  <%end if%>
			  
			  <%if str_Comunicacao = "AMB" then%>
				<option value="AMB" selected>Interno/Externa</option>
			  <%elseif Request("selComunicacao") = "AMB" then%>
				<option value="AMB" selected>Interno/Externa</option>
			  <%else%>
				<option value="AMB">Interno/Externa</option>
			  <%end if%> 		 
			</select>
		
		<%else
			'*** IMPRIME NA TELA O CONTEÚDO QUE SERIA MOSTRADO NO COMBO DE COMUNICAÇÃO
			if str_Comunicacao = "" then
				str_Comunicacao = Request("selComunicacao")
			end if
			
			if str_Comunicacao = "INT" then
				Response.write "Interno"
			else
				Response.write "Externo"
			end if
			
		end if%>
	  </td>
      <td width="3%">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td valign="top" class="campob"> <div align="right">O que comunicar:</div></td>
      <td>        
	  <table width="98%"  border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="46%" valign="top">
				<%if Request("txtOqueComunicar") <> "" then%>
					<textarea cols="34" rows="4" name="txtOqueComunicar" id="txtOqueComunicar" <%=boo_Escrita%> onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=Request("txtOqueComunicar")%></textarea>
				<%else%>
					<textarea cols="34" rows="4" name="txtOqueComunicar" id="txtOqueComunicar" <%=boo_Escrita%> onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=str_OqueComunicar%></textarea>
					<table width="100%"  height="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>				
						<td>
							<font face="Verdana" size="2"><font size="1">Caracteres digitados&nbsp;</font><b> 
							<input type="text" name="txttamanhoOqueComunicar" size="5" value="0" maxlength="50" readonly>
							</b></font><font face="Verdana" size="1">(Máximo 5000 caracteres)</font> 
						</td>													
					 </tr>
                    </table>
				<%end if%>					
			</td>
            <td width="54%">
              <table width="97%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
				  <%if (str_Acao = "A" or str_Acao = "C") and str_arquivo1 <> "" or str_arquivo2 <> "" or str_arquivo3 <> "" then%>
                  		<td width="107" class="campo"><div align="center" class="campob">Ver Arquivos:</div></td>
				  <%else%> 
				  	<td width="7">&nbsp;</td>
				  <%end if%>				  
                  <td width="255" class="campo"><span class="campob"><%if str_Acao <> "C" then%>Anexa arquivos:<%end if%></span></td>
                </tr>
                <tr>
				  <%if (str_Acao = "A" or str_Acao = "C")and str_arquivo1 <> "" then%>
                  		<td width="107"><div align="center"><a href="file:<%=str_arquivo1%>" target="_new"><img src="../img/anexa_arquivo_01.gif" width="20" height="20" border="0"></a><span class="campo"><a href="file:<%=str_arquivo1%>" target="_new" class="link">Arquivo 1</a></span></div></td>
				  <%else%> 
				  	<td>&nbsp;</td>
				  <%end if%>                  
                  <td>
				  	<%if str_Acao <> "C" then%>
						<%if Request("filArquivo1") <> "" then%>
							<input name="filArquivo1" type="file" value="<%=Request("filArquivo1")%>" class="campo" id="filArquivo1">
						<%else%>
							<input name="filArquivo1" type="file" class="campo" id="filArquivo1">
						<%end if%>		
					<%end if%>		  	
				  </td>
                </tr>
                <tr>
				  <%if (str_Acao = "A" or str_Acao = "C") and str_arquivo2 <> "" then%>
                  		<td width="107"><div align="center"><a href="file:<%=str_arquivo2%>" target="_new"><img src="../img/anexa_arquivo_01.gif" width="20" height="20" border="0"></a><span class="campo"><a href="file:<%=str_arquivo2%>" target="_new" class="link">Arquivo 2</a></span></div></td>
				  <%else%> 
				  	<td>&nbsp;</td>
				  <%end if%>                   
                  <td>
				  	<%if str_Acao <> "C" then%>
						<%if Request("filArquivo2") <> "" then%>
							<input name="filArquivo2" type="file" value="<%=Request("filArquivo2")%>" class="campo" id="filArquivo2">
						<%else%>
							<input name="filArquivo2" type="file" class="campo" id="filArquivo2">
						<%end if%>
					<%end if%>				  	
				  </td>				 
                </tr>
                <tr>
				  <%if (str_Acao = "A" or str_Acao = "C") and str_arquivo3 <> "" then%>
                  		<td width="107"><div align="center"><a href="file:<%=str_arquivo3%>" target="_new"><img src="../img/anexa_arquivo_01.gif" width="20" height="20" border="0"></a><span class="campo"><a href="file:<%=str_arquivo3%>" target="_new">Arquivo 3</a></span></div></td>
				  <%else%> 
				  	<td>&nbsp;</td>
				  <%end if%>                  
                  <td>
				  	<%if str_Acao <> "C" then%>
						<%if Request("filArquivo3") <> "" then%>
							<input name="filArquivo3" type="file" value="<%=Request("filArquivo3")%>" class="campo" id="filArquivo3">
						<%else%>
							<input name="filArquivo3" type="file" class="campo" id="filArquivo3">
						<%end if%>		
					<%end if%>			
				  </td>
                </tr>
                <tr>
					<%if str_Acao <> "C" then%>		
					    <td>&nbsp;</td>			    				
                  		<td colspan="2" align="left"><font face="Verdana" size="1">&nbsp;&nbsp;&nbsp;Os arquivos de anexo dever&atilde;o estar <br>&nbsp;&nbsp;&nbsp;em qualquer diret&oacute;rio no drive &quot;P&quot;</font></td>
                	<%end if%>
				</tr>
              </table></td>
          </tr>	
        </table></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Para quem comunicar:</div></td>
      	<td>
			<%if Request("txtAQuemComunicar") <> "" then%>
				<input name="txtAQuemComunicar" size="45" maxlength="50" type="text" <%=boo_Escrita%> class="txtCampo" value="<%=Request("txtAQuemComunicar")%>">
			<%else%>
				<input name="txtAQuemComunicar" size="45" maxlength="50" type="text" <%=boo_Escrita%> class="txtCampo" value="<%=str_AQuemComunicar%>">
			<%end if%>		  		
		</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Unidade &Oacute;rg&atilde;o:</div></td>
      <td>
	  	<%if Request("txtUnidadeOrgao") <> "" then%>
			<input name="txtUnidadeOrgao" size="45" maxlength="50" type="text" <%=boo_Escrita%> class="txtCampo" value="<%=Request("txtUnidadeOrgao")%>">
		<%else%>
			<input name="txtUnidadeOrgao" size="45" maxlength="50" type="text" <%=boo_Escrita%> class="txtCampo" value="<%=str_UnidadeOrgao%>">			
		<%end if%>  	
		</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Data Limite para Comunicar:</div></td>
      <td>
	  	<%if Request("txtQuandoOcorre") <> "" then%>
			<input name="txtQuandoOcorre" type="text" size="10" readonly maxlength="10" class="txtCampo" value="<%=Request("txtQuandoOcorre")%>">
		<%else%>
			<input name="txtQuandoOcorre" type="text" size="10" readonly maxlength="10" class="txtCampo" value="<%=str_txtQuandoOcorre%>">
		<%end if
		if str_Acao <> "C" then
		%> 	  	
	   	<a href="javascript:show_calendar(true,'frm_Plano_PCM.txtQuandoOcorre','DD/MM/YYYY')"><img  src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a><% end if %></td>
	  </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="4"><hr></td>
    </tr>
    <tr> 
      <td colspan="2"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Equipe de comunica&ccedil;&atilde;o</strong></font></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Respons&aacute;vel pelo Conteúdo:</div></td>
      <td>
	  	<%if Request("txtRespConteudo") <> "" then%>
			<input name="txtRespConteudo" size="45" maxlength="100" type="text" <%=boo_Escrita%> class="txtCampo" value="<%=Request("txtRespConteudo")%>">
		<%else%>
			<input name="txtRespConteudo" size="45" maxlength="100" type="text" <%=boo_Escrita%> class="txtCampo" value="<%=str_RespConteudo%>">
		<%end if%>	  
		</td>		
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Respons&aacute;vel pela Divulgação:</div></td>
      <td>
	  	<%if Request("txtRespDivulg") <> "" then%>
			<input name="txtRespDivulg" size="45" maxlength="100" type="text" <%=boo_Escrita%> class="txtCampo" value="<%=Request("txtRespDivulg")%>">
		<%else%>
			<input name="txtRespDivulg" size="45" maxlength="100" type="text" <%=boo_Escrita%> class="txtCampo" value="<%=str_RespDivulg%>">
		<%end if%>	  	
	  </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td class="campob"><div align="right">Como:</div></td>
      <td>
	  	<%if Request("txtComo") <> "" then%>
			<input name="txtComo" type="text" size="45" maxlength="50" <%=boo_Escrita%> class="txtCampo" value="<%=Request("txtComo")%>">
		<%else%>
			<input name="txtComo" type="text" size="45" maxlength="50" <%=boo_Escrita%> class="txtCampo" value="<%=str_Como%>">
		<%end if%>	  	
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
	if strTipoResponsavel <> "" then				
		if strCampo = "txtAprovadorPB" then
			strUsuaAprovadorPB 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 		
		end if				
	end if
	%>
	
	<%if strUsuaAprovadorPB <> "" then%>	
		<input type="hidden" value="<%=strUsuaAprovadorPB%>" name="hdUsuaAprovadorPB">
	<%else%>
		<input type="hidden" value="<%=Request("hdUsuaAprovadorPB")%>" name="hdUsuaAprovadorPB">
	<%end if%>	
	
    <tr> 
      <td class="campob"><div align="right">Aprovador PB:</div></td>
      <td class="campob">	
	  	<%if Request("txtAprovadorPB") <> "" then%>
			<input name="txtAprovadorPB" maxlength="4" type="text" <%=boo_Escrita%> size="5" value="<%=Request("txtAprovadorPB")%>">
		<%else%>	  	
			<input name="txtAprovadorPB" maxlength="4" type="text" <%=boo_Escrita%> size="5" value="<%=str_AprovadorPB%>">
		<%end if%>
		<a href="javascript:Localiza_Usuario('Legado','txtAprovadorPB');"><img src="../img/botao_localiza_Usuario.gif" border="0" ></a>				
		<%
		if strUsuaAprovadorPB <> "" then
			Response.write strUsuaAprovadorPB
		else
			Response.write Request("hdUsuaAprovadorPB") 
		end if
		%> 		
	  </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Data de Aprovação:</div></td>
      <td>
	  	<%if Request("txtDtAprovacao") <> "" then%>
			<input name="txtDtAprovacao" type="text" class="txtCampo" size="10" readonly maxlength="10" value="<%=Request("txtDtAprovacao")%>">
		<%else%>
			<input name="txtDtAprovacao" type="text" class="txtCampo" size="10" readonly maxlength="10" value="<%=str_txtDtAprovacao%>">
		<%end if
		if str_Acao <> "C" then
		%>       
		<a href="javascript:show_calendar(true,'frm_Plano_PCM.txtDtAprovacao','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a><% end if %> </td>
      <td>&nbsp;
	</td>			
      <td width="1%">&nbsp;</td>
    </tr>
    <tr>
      <td class="campob">&nbsp;</td>
      <td><input type="hidden" value="<%=int_Cd_ProjetoProject%>" name="pCdProjProject">
        <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
        <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
        <input type="hidden" value="<%=int_Plano2%>" name="pintPlano2">
        <input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
        <input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
        <input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">
        <input type="hidden" value="<%=str_Acao%>" name="pAcao">
        <input type="hidden" value="<%=str_CdSeqPCM%>" name="pCdSeqPCM">
        <input type="hidden" value="<%=int_CD_Onda%>" name="pOnda">
        <input type="hidden" value="PCM" name="pPlano">
        <input type="hidden" value="" name="pCampo">
        <input type="hidden" value="" name="pTipoResponsavel">
        <input type="hidden" value="" name="pChaveUsua"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
 
<table width="625" border="0" align="center">
  <tr>
    <td width="85" height="24"> <%
  if str_Acao <> "C" then
  %><a href="javascript:confirma_pcm()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a>
  <%
	end if
	%></td>
    <td width="24"><b></b></td>
    <td width="145">&nbsp;</td>
    <td width="193"><%if str_Acao = "C" then%>
      <a href="javascript:history.go(-1)"><img src="../img/botao_voltar.gif" width="85" height="19" border="0"></a>      <%end if%></td>
    <td width="9"></td>
    <td width="22"></td>
    <td width="10">&nbsp;</td>
    <td width="103"><div align="center"></div></td>
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
