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

if Request("selOnda") <> "" then
	int_CD_Onda		= Request("selOnda")
else
	int_CD_Onda		= Request("pOnda")
end if

intResData = request("pResData")

if Request("pPlano") <> "" then
	int_Plano = request("pPlano")
else
	int_Plano = request("pintPlano")
end if

if request("pIdAtividade") <> "" then 
	int_IdAtividade = request("pIdAtividade")
end if

'Response.write "Teste " & int_IdAtividade
'Response.end

strNomePlanoOrigem = request("pPlano_Origem")
int_CD_Onda = request("pOnda")

strPlanoOriginal = request("pPlanoOriginal")

'if request("selFases") <> "" then
'	str_Fase = request("selFases") 
'else
	str_Fase = request("pFase")
'end if

'Response.write "strPlanoOriginal " & strPlanoOriginal
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
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAC_TX_PROBLEMAS "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAC_TX_ACOES_CORR_CONT "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPAC_DT_APROVACAO "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_TRAT_PROC "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_SIN_TEC "				
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_SIN_FUN "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_TX_OPERACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_CD_NR_USUARIO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_DT_ATUALIZACAO "
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " FROM XPEP_PLANO_TAREFA_PAC"
		'str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PLTA_NR_SEQUENCIA_TAREFA = " & rds_sqlGeralAlteracao("PLTA_NR_SEQUENCIA_TAREFA")
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PLTA_NR_SEQUENCIA_TAREFA = " & int_Id_TarefaProject
						
		'Response.write str_sqlAtividadeAlt
		'Response.end										
		Set rds_sqlAtividadeAlt = db_Cogest.Execute(str_sqlAtividadeAlt)	
		
		rds_sqlGeralAlteracao.close
		set rds_sqlGeralAlteracao = nothing
		
		if not rds_sqlAtividadeAlt.eof then					
			
			str_txtProblemas  			= Trim(rds_sqlAtividadeAlt("PPAC_TX_PROBLEMAS"))
			str_txtAcoesCorrConting 	= Trim(rds_sqlAtividadeAlt("PPAC_TX_ACOES_CORR_CONT"))			
			str_UsuarioResponsavel		= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_TRAT_PROC"))	
			str_txtRespSinergiaTec		= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_SIN_TEC"))		
			str_txtRespSinergiaFunc		= Trim(rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_SIN_FUN"))			
																					
			strDia = ""		
			strMes = ""
			strAno = ""
			vetDTAprovacao_PAC = split(Trim(rds_sqlAtividadeAlt("PPAC_DT_APROVACAO")),"/")							
			strDia = trim(vetDTAprovacao_PAC(1))
			if cint(strDia) < 10 then
				strDia = "0" & strDia
			end if			
			strMes = trim(vetDTAprovacao_PAC(0))			
			if cint(strMes) < 10 then
				strMes = "0" & strMes
			end if
			strAno = trim(vetDTAprovacao_PAC(2))
			str_txtDTAprovacao_PAC = strDia & "/" & strMes & "/" & strAno 			
		end if	
		rds_sqlAtividadeAlt.close
		set rds_sqlAtividadeAlt = nothing
	end if    
end if

strReadOnly = ""
if str_Acao ="C" then
	strReadOnly = "readonly"
	str_Texto_Acao = "Consulta"
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

str_Atividade_Origem = strNomePlanoOrigem & " - " & str_NomeAtividade
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>:: Cutover - Plano PAC</title>
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
	<script src="../js/troca_lista.js" language="javascript"></script>
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
					
					if (strobjnome=='txtAtividade')
					{
						document.forms[0].txtAtividade.value = strvalor.substr(0,i);
					}
					else
					{
						document.forms[0].txtAcoesCorrConting.value = strvalor.substr(0,i);
					}
					break;
				}
			}
		}		
	}
				
	function confirma_pac()
	{
		var txt_Problemas			= document.frm_Plano_PAC.txtProblemas.value;	
		var txt_AcoesCorrConting 	= document.frm_Plano_PAC.txtAcoesCorrConting.value;		 
		var txt_DTAprovacao_PAC 	= document.frm_Plano_PAC.txtDTAprovacao_PAC.value;	
		var txt_UsuarioResponsavel 	= document.frm_Plano_PAC.txtUsuarioResponsavel.value; 			
		var txt_RespTecSinGeral		= document.frm_Plano_PAC.txtRespTecSinGeral.value; 					
		var txt_RespFunSinGeral 	= document.frm_Plano_PAC.txtRespFunSinGeral.value; 			
								
		//*** Ações Corretivas/Contingências		  
		if(txt_Problemas == "")
		  {
		  alert("É obrigatório o preenchimento do campo Problemas!");
		  document.frm_Plano_PAC.txtProblemas.focus();
		  return;
		  }
		  
		//*** Ações Corretivas/Contingências		  
		if(txt_AcoesCorrConting == "")
		  {
		  alert("É obrigatório o preenchimento do campo Ações Corretivas/Contingências!");
		  document.frm_Plano_PAC.txtAcoesCorrConting.focus();
		  return;
		  }
					
		//*** Data da Aprovação	  
		//if(txt_DTAprovacao_PAC == "")
		//  {
		//  alert("É obrigatório o preenchimento do campo Data da Aprovação!");
		//  document.frm_Plano_PAC.txtDTAprovacao_PAC.focus();
		//  return;
		//  }
	    //else
		 //{
			//validaData(txt_DTAprovacao_PAC,'txtDTAprovacao_PAC','Data da Aprovação');
			//if (blnData) return; 	
		 //}
				  
		//*** Usuário Responsável  
		if(txt_UsuarioResponsavel == '')
		  {
		  alert("É obrigatório o preenchimento do campo Usuário Responsável!");
		  document.frm_Plano_PAC.txtUsuarioResponsavel.focus();
		  return;
		  } 	
					
		//*** Responsável Técnico - Sinergia 
		if(txt_RespTecSinGeral == '')
		  {
		  alert("É obrigatório o preenchimento do campo Responsável Sinergia - Técnico!");
		  document.frm_Plano_PAC.txtRespTecSinGeral.focus();
		  return;
		  } 		
		  
		//*** Responsável Sinergia - Funcional
		if(txt_RespFunSinGeral == '')
		  {
		  alert("É obrigatório o preenchimento do campo Responsável Sinergia - Funcional!");
		  document.frm_Plano_PAC.txtRespFunSinGeral.focus();
		  return;
		  } 		
		  
	   document.frm_Plano_PAC.action="grava_plano.asp?pPlano=PAC";           
	   document.frm_Plano_PAC.submit();				
	}	

	function envia_Email()
	{									
		document.frm_Plano_PAC.action="envia_email.asp?pPlano=PAC";           
   		document.frm_Plano_PAC.submit();				
	}
	
	function confirma_Exclusao()
	{
		  if(confirm("Confirma a exclusão deste Registro?"))
		  {
		    document.frm_Plano_PAC.pAcao.value = 'E';			
			document.frm_Plano_PAC.action='grava_plano.asp?pPlano=PAC' 			        
			document.frm_Plano_PAC.submit();
		  }
	}	
	
	function Localiza_Usuario(strTipoResponsavel,strCampo)
	{
		if (strCampo == 'txtUsuarioResponsavel')
		{
			strUsuario = document.frm_Plano_PAC.txtUsuarioResponsavel.value;		
		
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Responsável pelo Tratamento do Procedimento!");
				document.frm_Plano_PAC.txtUsuarioResponsavel.focus();
				return;
			}
		}
		
		if (strCampo == 'txtRespTecSinGeral')
		{
			strUsuario = document.frm_Plano_PAC.txtRespTecSinGeral.value;		
		
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Responsável Sinergia - Técnico!");
				document.frm_Plano_PAC.txtRespTecSinGeral.focus();
				return;
			}
		}					
				
		if (strCampo == 'txtRespFunSinGeral')
		{
			strUsuario = document.frm_Plano_PAC.txtRespFunSinGeral.value;		
		
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Responsável Sinergia - Funcional!");
				document.frm_Plano_PAC.txtRespFunSinGeral.focus();
				return;
			}
		}			
				
		document.frm_Plano_PAC.pTipoResponsavel.value = strTipoResponsavel;	
		document.frm_Plano_PAC.pChaveUsua.value = strUsuario.toUpperCase();	
						
		document.frm_Plano_PAC.action='inclui_altera_plano_pac.asp?pTipoResponsavel=' + strTipoResponsavel + '&pCampo=' + strCampo;
		document.frm_Plano_PAC.submit();			
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
    <td width="5%">&nbsp;</td>
    <td width="92%" class="subtitulo"><table width="92%" border="0" cellpadding="0" cellspacing="7">
		<% if str_Atividade_Origem <> "" then %>
        <tr>
          <td><div align="right"><span class="subtitulob">Atividade Origem:</span></div></td>
          <td class="subtitulo"><%=str_Atividade_Origem%></td>
          <td class="subtitulo"><div align="center" class="campo">A&ccedil;&atilde;o</div></td>
        </tr>
		<% end if %>
        <tr>
          <td>&nbsp;</td>
          <td class="subtitulo">&nbsp;</td>
          <td width="85" bgcolor="#EFEFEF"><div align="center"><span class="campob"><%=str_Texto_Acao%></span></div></td>
        </tr>
        <tr>
          <td width="165"><div align="right" class="subtitulob">Onda:</div></td>
          <td class="subtitulo"><%=str_Desc_Onda%></td>
          <td class="subtitulo">&nbsp;</td>
        </tr>
        <tr>
          <td><div align="right"><span class="subtitulob">Plano:</span></div></td>
          <td width="574" class="subtitulo">Plano de A&ccedil;&otilde;es Corretivas e Conting&ecirc;ncias - PAC</td>
          <td width="85" class="subtitulo">&nbsp;</td>
        </tr>
      </table>      </td>
    <td width="3%">&nbsp;</td>
  </tr>
</table>
<table width="75%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr> 
    <td width="17%" bgcolor="#EEEEEE"> <div align="right" class="campo">Atividade:</div></td>
    <td colspan="3" class="campob"><%=str_NomeAtividade%></td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Respons&aacute;vel:</div></td>
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
    <td width="20%" bgcolor="#EEEEEE"> <div align="right" class="campo">Data de 
        T&eacute;rmino:</div></td>
    <td width="33%" class="campob"><%=dat_Dt_Termino%></td>
  </tr>
</table>
<hr>
<form name="frm_Plano_PAC" method="post" action="">
  <table width="98%" border="0">
    <tr>
      <td class="campo">&nbsp;</td>
      <td height="25" valign="top" class="campob"><div align="right">Problemas:</div></td>
      <td class="campo">
	  <%if Request("txtProblemas") <> "" then%>
	  		<textarea name="txtProblemas" cols="34" rows="4" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);" <%=strReadOnly%>><%=Request("txtProblemas")%></textarea>
      <%else%>
			<textarea name="txtProblemas" cols="34" rows="4" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);" <%=strReadOnly%>><%=str_txtProblemas%></textarea>
	  <%end if%>
	  </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="1%" class="campo">&nbsp;</td>
      <td width="20%" height="25" valign="top" class="campob"><div align="right">A&ccedil;&otilde;es 
          Corretivas/Conting&ecirc;ncias:</div></td>
      <td width="76%" class="campo">
	   	<%if Request("txtAcoesCorrConting") <> "" then%>
	  		<textarea name="txtAcoesCorrConting" cols="34" rows="4" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);" <%=strReadOnly%>><%=Request("txtAcoesCorrConting")%></textarea>
      	<%else%>
			<textarea name="txtAcoesCorrConting" cols="34" rows="4" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);" <%=strReadOnly%>><%=str_txtAcoesCorrConting%></textarea>
	  	<%end if%>
	  </td>
      <td width="2%">&nbsp;</td>
      <td width="1%">&nbsp;</td>
    </tr>
    <tr> 
      <td class="campo">&nbsp;</td>
      <td height="25" class="campob"><div align="right">Data da Aprova&ccedil;&atilde;o:</div></td>
      <td class="campo"> 
	  	<%if Request("txtDTAprovacao_PAC") <> "" then%>
	  		<input type="text" name="txtDTAprovacao_PAC" class="txtCampo" size="10" value="<%=Request("txtDTAprovacao_PAC")%>" readonly> 
      	<%else%>
			<input type="text" name="txtDTAprovacao_PAC" class="txtCampo" size="10" value="<%=str_txtDTAprovacao_PAC%>" readonly> 
	  	<%end if	 
		if str_Acao <> "C" then%> 		
			<a href="javascript:show_calendar(true,'frm_Plano_PAC.txtDTAprovacao_PAC','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a>
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
		if strCampo = "txtUsuarioResponsavel" then
			strUsuarioRespProced = " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 	
		elseif strCampo = "txtRespTecSinGeral"  then
			strRespSinergiaTec   = " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 	
		elseif strCampo = "txtRespFunSinGeral" then
			strRespSinergiaFunc  = " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel)		
		end if				
	end if
	%>	
	
	<%if strUsuarioRespProced <> "" then%>	
		<input type="hidden" value="<%=strUsuarioRespProced%>" name="hdUsuarioRespProced">
	<%else%>
		<input type="hidden" value="<%=Request("hdUsuarioRespProced")%>" name="hdUsuarioRespProced">
	<%end if%>
	
	<%if strRespSinergiaTec <> "" then%>	
		<input type="hidden" value="<%=strRespSinergiaTec%>" name="hdRespSinergiaTec">
	<%else%>
		<input type="hidden" value="<%=Request("hdRespSinergiaTec")%>" name="hdRespSinergiaTec">
	<%end if%>	
	
	<%if strRespSinergiaFunc <> "" then%>	
		<input type="hidden" value="<%=strRespSinergiaFunc%>" name="hdRespSinergiaFunc">
	<%else%>
		<input type="hidden" value="<%=Request("hdRespSinergiaFunc")%>" name="hdRespSinergiaFunc">
	<%end if%>	
	
    <td class="campo">&nbsp;</td>
    <td class="campob"><div align="right">Respons&aacute;vel pelo Tratamento do 
        Procedimento:</div></td>
    <td class="campob" valign="bottom">	
		<%if Request("txtUsuarioResponsavel") <> "" then%>
			<input type="text" maxlength="4" value="<%=Request("txtUsuarioResponsavel")%>" name="txtUsuarioResponsavel" <%=strReadOnly%> size="5">
		<%else%>
			<input type="text" maxlength="4" value="<%=str_UsuarioResponsavel%>" name="txtUsuarioResponsavel" <%=strReadOnly%> size="5">
		<%end if%>
		<a href="javascript:Localiza_Usuario('Legado','txtUsuarioResponsavel');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>						
		<%
		if strUsuarioRespProced <> "" then
			Response.write strUsuarioRespProced
		else
			Response.write Request("hdUsuarioRespProced") 
		end if
		%>
    </td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="5">
	  <table width="88%" border="0">
		  <tr> 
			<td height="27" colspan="3"> <table width="58%" border="0">
				<tr> 
				  <td width="19%">&nbsp;</td>
				  <td width="81%" class="campob">Respons&aacute;vel Sinergia</td>
				</tr>
			  </table></td>
			<td width="29%">&nbsp;</td>
			<td width="29%">&nbsp;</td>
		  </tr>
		  <tr> 
			<td width="9%" valign="top" class="campo">&nbsp;</td>
			<td width="15%" valign="top"> <div align="right"> 
				 <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr> 
					<td height="25"> <div align="right"><span class="campob">T&eacute;cnico:</span>:</div></td>
				  </tr>
				</table>
			  </div></td>
			<td colspan="3" class="campob">
				<%if Request("txtRespTecSinGeral") <> "" then%>
                <input type="text" maxlength="4" value="<%=Request("txtRespTecSinGeral")%>" name="txtRespTecSinGeral" <%=strReadOnly%> size="5">				
              <%else%>
					<input type="text" maxlength="4" value="<%=str_txtRespSinergiaTec%>" name="txtRespTecSinGeral" <%=strReadOnly%> size="5">
				<%end if%>
				<a href="javascript:Localiza_Usuario('Sinergia','txtRespTecSinGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>						
				<%
				if strRespSinergiaTec <> "" then
					Response.write strRespSinergiaTec
				else
					Response.write Request("hdRespSinergiaTec") 
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
					<input type="text" maxlength="4" value="<%=Request("txtRespFunSinGeral")%>" name="txtRespFunSinGeral" <%=strReadOnly%> size="5">					
				<%else%>
					<input type="text" maxlength="4" value="<%=str_txtRespSinergiaFunc%>" name="txtRespFunSinGeral" <%=strReadOnly%> size="5">					
				<%end if%>
				<a href="javascript:Localiza_Usuario('Sinergia','txtRespFunSinGeral');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>						
				<%
				if strRespSinergiaFunc <> "" then
					Response.write strRespSinergiaFunc
				else
					Response.write Request("hdRespSinergiaFunc") 
				end if
				%>				
			</td>
		  </tr>
		</table>	  
	  </td>
    </tr>
	  <tr>     
      <td colspan="5">
	  	  <input type="hidden" value="<%=int_Cd_ProjetoProject%>" name="pCdProjProject">
	  	  <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
		  <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
		  <input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
		  <input type="hidden" value="<%=int_IdAtividade%>" name="pIdAtividade">
		  <input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
		  <input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">
		  <input type="hidden" value="<%=str_Acao%>" name="pAcao">
		  <input type="hidden" value="<%=int_CD_Onda%>" name="pOnda">
		  <input type="hidden" value="<%=str_Fase%>" name="pFase">
		  <input type="hidden" value="<%=str_Atividade_Origem%>" name="pPlanoOrigem">	
		  <input type="hidden" value="<%=str_Desc_Onda%>" name="pNomeOnda">	
		  <input type="hidden" value="" name="pTipoResponsavel">
		  <input type="hidden" value="" name="pChaveUsua">		 		  
	  </td>
    </tr>    
  </table>
  <table width="661" border="0" align="center">
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
      <td width="131" height="37">
	  	<%if str_Acao <> "C" then%>
	  		<a href="javascript:confirma_pac()"><img src="../img/enviar_01.gif" alt=":: Envia para confirmar operação" width="85" height="19" border="0"></a>			
		<%end if%>
	  </td>		
      <td width="9" height="37"><b></b></td>
      <td width="131" height="37">
	  	<%if str_Acao = "A" then%>
	  		<a href="javascript:envia_Email();"><img src="../img/botao_enviar_email.gif" alt=":: Encaminha email" border="0"></a>
   	    <%end if%>	 
	  </td>
	  <td width="6" height="37"></td>
      <td width="134" height="37">
	  	<%if str_Acao = "A" then%>
	  		<a href="javascript:confirma_Exclusao();"><img src="../img/botao_excluir.gif" alt=":: Exclui registro" border="0"></a>
     	<%end if%>	  </td>	  	
      <td width="8" height="37"></td>
      <td width="150" height="37"><%if str_Acao = "C" then%>
        <a href="javascript:history.go(-1)"><img src="../img/botao_voltar.gif" alt=":: Volta tela anterior" width="85" height="19" border="0"></a>
      <%end if%></td>
      <td width="58" height="37"><div align="center"></div></td>
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
		<p width="780" align="center"><img src="../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>
