<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Expires=0

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

if Request("pPlano") <> "" then
	int_Plano = request("pPlano")
else
	int_Plano = request("pintPlano")
end if

if Request("pTArefa") <> "" then
	int_Id_TarefaProject = Request("pTArefa")
else	
	int_Id_TarefaProject = Request("idTaskProject")
end if

int_CD_Onda = request("pOnda")
int_ResData = request("pResData")

str_Fase = request("pFase")
strPlanoOriginal = Request("pPlanoOriginal")

'response.Write(int_Cd_ProjetoProject & "<P>")
'response.Write(int_Plano & "<P>")
'response.Write(int_Id_TarefaProject & "<P>")
'response.Write(int_CD_Onda & "<P>")
'response.Write(int_ResData & "<P>")
'response.End()

str_rdbTpDesligamento = ""
int_selRespTecLeg = ""
int_selRespFunLeg = ""
str_txtGerTecRespLeg = ""
int_Cd_Sistema_Legado = ""

'QUANDO VEM DA GRAVAÇÃO DA FUNCIONALIDADE EU DEVO ENCONTRAR O COD DO 
IF int_Cd_ProjetoProject = "" then
	str_Sql_Cd_Projeto_Project = ""
	str_Sql_Cd_Projeto_Project = str_Sql_Cd_Projeto_Project & " SELECT "
	str_Sql_Cd_Projeto_Project = str_Sql_Cd_Projeto_Project & " PLAN_NR_SEQUENCIA_PLANO"
	str_Sql_Cd_Projeto_Project = str_Sql_Cd_Projeto_Project & " , PLAN_NR_CD_PROJETO_PROJECT"
	str_Sql_Cd_Projeto_Project = str_Sql_Cd_Projeto_Project & " , PLAN_NR_CD_ONDA"
	str_Sql_Cd_Projeto_Project = str_Sql_Cd_Projeto_Project & " FROM XPEP_PLANO_ENT_PRODUCAO"
	str_Sql_Cd_Projeto_Project = str_Sql_Cd_Projeto_Project & " WHERE  PLAN_NR_SEQUENCIA_PLANO = " & int_Plano
	Set rdsCd_Projeto_Project = db_Cogest.Execute(str_Sql_Cd_Projeto_Project)	
	if not rdsCd_Projeto_Project.EOF then
       int_Cd_ProjetoProject = rdsCd_Projeto_Project("PLAN_NR_CD_PROJETO_PROJECT")
	   int_CD_Onda = rdsCd_Projeto_Project("PLAN_NR_CD_ONDA")
	end if
	rdsCd_Projeto_Project.close
	set rdsCd_Projeto_Project = Nothing	
end if

if str_Acao = "I" then
   	str_Texto_Acao = "Inclusão"   
else
   	str_Texto_Acao = "Alteração"   
	
  	str_sqlGeralAlteracao = ""
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & "SELECT PLTA_NR_SEQUENCIA_TAREFA"
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & " FROM XPEP_PLANO_TAREFA_GERAL"
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & " WHERE  PLAN_NR_SEQUENCIA_PLANO = " & int_Plano
	str_sqlGeralAlteracao = str_sqlGeralAlteracao & " AND PLTA_NR_ID_TAREFA_PROJECT = " & int_Id_TarefaProject
	Set rds_sqlGeralAlteracao = db_Cogest.Execute(str_sqlGeralAlteracao)				
	
	if not rds_sqlGeralAlteracao.eof then		
		str_sqlAtividadeAlt = ""
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & "SELECT "		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & "  PLAN_NR_SEQUENCIA_PLANO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PLTA_NR_SEQUENCIA_TAREFA"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", SIST_NR_SEQUENCIAL_SISTEMA_LEGADO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPDS_TX_TIPO_DESLIGAMENTO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", PPDS_TX_GER_TEC_RESP_LEGADO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_LEG_TEC"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", USUA_CD_USUARIO_RESP_LEG_FUN"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_TX_OPERACAO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_CD_NR_USUARIO"
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & ", ATUA_DT_ATUALIZACAO"		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " FROM XPEP_PLANO_TAREFA_PDS"		
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & int_Plano
		str_sqlAtividadeAlt = str_sqlAtividadeAlt & " AND PLTA_NR_SEQUENCIA_TAREFA = " & rds_sqlGeralAlteracao("PLTA_NR_SEQUENCIA_TAREFA")
		'Response.write str_sqlAtividadeAlt
		'Response.end
		Set rds_sqlAtividadeAlt = db_Cogest.Execute(str_sqlAtividadeAlt)	
		
		if not rds_sqlAtividadeAlt.eof then				
		    int_Cd_Sistema_Legado = rds_sqlAtividadeAlt("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO")	
			str_rdbTpDesligamento = rds_sqlAtividadeAlt("PPDS_TX_TIPO_DESLIGAMENTO")
			str_selRespTecLeg = rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_LEG_TEC")
			str_selRespFunLeg = rds_sqlAtividadeAlt("USUA_CD_USUARIO_RESP_LEG_FUN")
			str_txtGerTecRespLeg = rds_sqlAtividadeAlt("PPDS_TX_GER_TEC_RESP_LEGADO")
		end if
		rds_sqlAtividadeAlt.close
		set rds_sqlAtividadeAlt = nothing	
	end if
	rds_sqlGeralAlteracao.close
	set rds_sqlGeralAlteracao = nothing		
end if

'response.Write(int_Cd_Sistema_Legado)
'response.Write(str_rdbTpDesligamento)
'response.Write(int_selRespTecLeg)
'response.Write(int_selRespFunLeg)
'response.Write(str_txtGerTecRespLeg)
'response.End()

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
'Response.write str_Sql_DadosAdicionais_Tarefa
'Response.end
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

str_Pds_Func = ""
str_Pds_Func = str_Pds_Func & " SELECT  "
str_Pds_Func = str_Pds_Func & " PPDS_NR_SEQUENCIA_FUNC"
str_Pds_Func = str_Pds_Func & " , PPDS_TX_FUNC_DESATIVADAS"
str_Pds_Func = str_Pds_Func & " , PPDS_DT_DESLIGAMENTO"
str_Pds_Func = str_Pds_Func & " , PPDS_TX_HR_DESLIGAMENTO"
str_Pds_Func = str_Pds_Func & " , PPDS_TX_PROC_DESLIGAMENTO"
str_Pds_Func = str_Pds_Func & " , PPDS_TX_DEST_DD_TEMPO_RETENCAO"
str_Pds_Func = str_Pds_Func & " FROM XPEP_PLANO_TAREFA_PDS_FUNC"
str_Pds_Func = str_Pds_Func & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & int_Plano  
str_Pds_Func = str_Pds_Func & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_Id_TarefaProject
'response.Write(str_Pds_Func)
'response.End()

strReadOnly = ""
if str_Acao = "C" then
	strReadOnly = "readonly"
	strDisabled = "disabled"
	str_Texto_Acao = "Consulta"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>:: Cutover - Plano PDS</title>
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
		
	function confirma_pds()
	{		

       if (document.frm_Plano_PDS.selSistLegado.selectedIndex == 0)
	      {
		  alert("É obrigatório a seleção do campo Sistema Legado pelo Legado!");
		  document.frm_Plano_PDS.selSistLegado.focus();
		  return;
		  }

        //alert(document.frm_Plano_PDS.selRespTecLeg.selectedIndex)
		//*** Descrição da Parada	
		if ((!document.frm_Plano_PDS.rdbTpDesligamento[0].checked)&&
		   (!document.frm_Plano_PDS.rdbTpDesligamento[1].checked))
		  {
		  alert("É obrigatório o preenchimento do campo Tipo de desligamento!");
		  document.frm_Plano_PDS.rdbTpDesligamento.focus();
		  return;
		  }

       if (document.frm_Plano_PDS.txtRespTecLeg.value == '')
	      {
		  alert("É obrigatório o preenchimento do campo Responsável Legado - Técnico!");
		  document.frm_Plano_PDS.txtRespTecLeg.focus();
		  return;
		  }

       if (document.frm_Plano_PDS.txtRespFunLeg.value == '')
	      {
		  alert("É obrigatório o preenchimento do campo Responsável Legado - Funcional!");
		  document.frm_Plano_PDS.txtRespFunLeg.focus();
		  return;
		  }		
	
       if (document.frm_Plano_PDS.txtGerTecRespLeg.value == "")
	      {
		  alert("É obrigatório o preenchimento do campo Gerência responsável!");
		  document.frm_Plano_PDS.txtGerTecRespLeg.focus();
		  return;
		  }

	   document.frm_Plano_PDS.action="grava_plano.asp?pPlano=PDS";           
	   document.frm_Plano_PDS.submit();				
	}	
	
	function confirma_Exclusao_Parcial()
	{
		  if(confirm("Confirma a exclusão desta Funcionalidade ?"))
		  {
		    document.frm_Plano_PDS_Func.pAcao2.value="E";
			document.frm_Plano_PDS_Func.action="grava_sub_ativ.asp?pPlano=PDS";           
			document.frm_Plano_PDS_Func.submit();				
		  }
	}	
	
	function Localiza_Usuario(strTipoResponsavel,strCampo)
	{	
		if (strCampo == 'txtRespTecLeg')
		{
			strUsuario = document.frm_Plano_PDS.txtRespTecLeg.value;
			
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Responsável Legado - Técnico!");
				document.frm_Plano_PDS.txtRespTecLeg.focus();
				return;
			}		
		}
				
		if (strCampo == 'txtRespFunLeg')
		{
			strUsuario = document.frm_Plano_PDS.txtRespFunLeg.value;
			
			if (strUsuario == '')
			{			
				alert("É obrigatório o preenchimento do campo Responsável Legado - Funcional!");
				document.frm_Plano_PDS.txtRespFunLeg.focus();
				return;
			}		
		}
		
		document.frm_Plano_PDS.pTipoResponsavel.value = strTipoResponsavel;	
		document.frm_Plano_PDS.pChaveUsua.value = strUsuario.toUpperCase();	
		document.frm_Plano_PDS.pCampo.value = strCampo;						
		
		document.frm_Plano_PDS.action='inclui_altera_plano_pds.asp?pTipoResponsavel=' + strTipoResponsavel + '&pCampo=' + strCampo;
		document.frm_Plano_PDS.submit();			
	}
	
	function confirma_Exclusao_Total()
		{
			  if(confirm("Confirma a exclusão deste Registro?"))
			  {
				document.frm_Plano_PDS.pAcao.value = 'E';			
				document.frm_Plano_PDS.action='grava_plano.asp?pPlano=PDS' 			        
				document.frm_Plano_PDS.submit();
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
    <table width="87%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="6%">&nbsp;</td>
    <td width="81%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo"><table width="99%" border="0" cellpadding="0" cellspacing="7">
        <tr> 
          <td width="29%"><div align="right" class="subtitulob">Onda:</div></td>
          <td width="71%" class="subtitulo"><%=str_Desc_Onda%></td>
        </tr>
        <tr> 
          <td><div align="right"><span class="subtitulob">Plano:</span></div></td>
          <td class="subtitulo">Plano de Desligamento de Sistemas Legados - PDS </td>
        </tr>
      </table>      </td>
    <td><table width="100%"  border="0">
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
    <td width="17%" bgcolor="#EEEEEE"> <div align="right" class="campo">Linha:</div></td>
    <td colspan="3"><span class="campob"><%=str_NomeAtividade%></span></td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Respons&aacute;vel:</div></td>
    <td colspan="3"><span class="campob"><%=str_Nome_Responsavel%></span></td>
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
    <td width="21%"><span class="campob"><%=dat_Dt_Inicio%></span></td>
    <td width="20%" bgcolor="#EEEEEE"><div align="right" class="campo">Data de T&eacute;rmino:</div></td>
    <td width="33%"><span class="campob"><%=dat_Dt_Termino%></span></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td height="3" bgcolor="#CCCCCC"></td>
  </tr>
</table>
<form name="frm_Plano_PDS" method="post">
  <table width="100%"  border="0" cellspacing="5" cellpadding="0">
    <tr>
      <td class="campob">Sistema Legado: </td>
      <td><% intIndice="selSistLegado"
             int_Cd_Usuario = int_selRespTecLeg			 
			 if Request("selSistLegado") <> "" then			 		 
			 	int_Cd_Sistema_Legado = Request("selSistLegado")
			 else
			    if str_Acao = "I" then
			       int_Cd_Sistema_Legado = 0
				end if   
			 end if 
			 'response.Write("  Sistema  " & int_Cd_Sistema_Legado)
			 %>        
        <!--#include file="../includes/inc_combo_Sistema_Legado_Geral.asp" --></td></tr>
    <tr>
      <td width="19%" class="campob">Tipo de desligamento: </td>
      <td width="81%">
	  <%if str_rdbTpDesligamento = "1" then
	        str_Marca_1 = "checked"
			str_Marca_2 = ""
		 elseIf str_rdbTpDesligamento = "2" then
  		    str_Marca_1 = ""
		    str_Marca_2 = "checked"
	     elseif Request("rdbTpDesligamento") = "1" then
		 	str_Marca_1 = "checked"
			str_Marca_2 = ""
		 elseif Request("rdbTpDesligamento") = "2" then
		 	str_Marca_1 = ""
			str_Marca_2 = "checked"			 
		 else		 
	        str_Marca_1 = ""
		    str_Marca_2 = ""
	     end if
      %>
	  <table width="200">
        <tr>
          <td><label>
            <input name="rdbTpDesligamento" type="radio" <%=strDisabled%> value="1" <%=str_Marca_1%>>
            <span class="campo">Total</span></label></td>
          <td><label>
            <input type="radio" name="rdbTpDesligamento" <%=strDisabled%> value="2" <%=str_Marca_2%>>
            <span class="campo">Parcial</span></label></td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td><span class="campob">Respons&aacute;vel Legado</span></td>
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
		if strCampo = "txtRespTecLeg" then
			strRespTecLegado 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 		
		elseif strCampo = "txtRespFunLeg" then
			strRespFunLegado 	= " - " & RetornaNomeUsuario(strChaveUsuario, strTipoResponsavel) 		
		end if					
	end if
	%>
	
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
	
    <tr>
      <td class="campob"><div align="right">T&eacute;cnico:</div></td>
      <td class="campob">
	  	<%if Request("txtRespTecLeg") <> "" then%>
			<input name="txtRespTecLeg" type="text" class="txtCampo" <%=strReadOnly%> value="<%=Request("txtRespTecLeg")%>" size="5" maxlength="4">	  	
		<%else%>	  	
			<input name="txtRespTecLeg" type="text" class="txtCampo" <%=strReadOnly%> value="<%=str_selRespTecLeg%>" size="5" maxlength="4">	  	
		<%end if%>
		<a href="javascript:Localiza_Usuario('Legado','txtRespTecLeg');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
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
      <td class="campob"><div align="right">Funcional:</div></td>
      <td class="campob">
	  	<%if Request("txtRespFunLeg") <> "" then%>
			<input name="txtRespFunLeg" type="text" class="txtCampo" <%=strReadOnly%> value="<%=Request("txtRespFunLeg")%>" size="5" maxlength="4">
		<%else%>	  	
			<input name="txtRespFunLeg" type="text" class="txtCampo" <%=strReadOnly%> value="<%=str_selRespFunLeg%>" size="5" maxlength="4">
		<%end if%>
		<a href="javascript:Localiza_Usuario('Legado','txtRespFunLeg');"><img src="../img/botao_localiza_Usuario.gif" border="0"></a>				
		<%
		if strRespFunLegado <> "" then
			Response.write strRespFunLegado
		else
			Response.write Request("hdRespFunLegado") 
		end if
		%>	
	  </td>
    </tr>
    <tr>
      <td class="campob">Ger&ecirc;ncia T&eacute;cnica Respons&aacute;vel pelo Legado:</td>
      <td>
	  	<%if Request("txtGerTecRespLeg") <> "" then%>
	  		<input name="txtGerTecRespLeg" type="text" <%=strReadOnly%> class="txtCampo" value="<%=Request("txtGerTecRespLeg")%>" size="70" maxlength="100"></td>
    	<%else%>
			<input name="txtGerTecRespLeg" type="text"<%=strReadOnly%>  class="txtCampo" value="<%=str_txtGerTecRespLeg%>" size="70" maxlength="100"></td>
		<%end if%>
	</tr>
    <tr>
      <td class="campob">&nbsp;</td>
      <td>
	    <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
        <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
        <input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
        <input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
        <input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">
        <input type="hidden" value="<%=str_Acao%>" name="pAcao">		
		<input type="hidden" value="" name="pCampo">
		<input type="hidden" value="" name="pTipoResponsavel">
		<input type="hidden" value="" name="pChaveUsua">
		</td>	
    </tr>
  </table>
  <%if str_Acao <> "C" then%>  
	  <%if str_Acao = "A" then%>
		  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
			<tr> 
				  <td width="1%">&nbsp;</td>
				   <td colspan="3" align="left" valign="bottom">
					<div class="campob">
					  <table width="48%" border="0">
						<tr>
						  <td width="95%">Link com Plano de A&ccedil;&otilde;es Corretivas / Conting&ecirc;ncia (PAC):</td>
						  <td width="5%"><div class="campob"><a href="encaminha_plano.asp?selTipoCadastro=PAC&pSiglaPlano=PAC&pAtividade_Origen=<%="PSD - " & str_NomeAtividade%>&selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Plano%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></div></td>
						</tr>
					  </table>
					</div>
					</td>	      
				  <td width="2%">&nbsp;</td>
				</tr> 
			</table>
		<%end if%>
	
	  <table width="625" border="0" align="center">
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td align="center">&nbsp;</td>
		  <td>&nbsp;</td>
		  <td></td>
		  <td></td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		</tr>
		<tr>
		  <td width="85"><a href="javascript:confirma_pds()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
		  <td width="22"><b>
			<input type="hidden" value="" name="pTipoResponsavel2">
		  </b></td>
		  <td width="149"> 
			<%if str_Acao = "A" then%>
				<a href="javascript:confirma_Exclusao_Total();"><img src="../img/botao_excluir.gif" width="85" height="19" border="0"></a>
			<%end if%>
		  </td>
		  <td width="204"><%if str_Acao = "C" then%>
            <a href="javascript:history.go(-1)"><img src="../img/botao_voltar.gif" width="85" height="19" border="0"></a>
            <%end if%></td>
		  <td width="9"></td>
		  <td width="8"></td>
		  <td width="7">&nbsp;</td>
		  <td width="107"><div align="center"></div></td>
		</tr>
	  </table>
  <%end if
  
  if str_Acao = "A" or str_Acao = "C" then%>  
  </form>
<form name="frm_Plano_PDS_Func" method="post" action="">
  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr><td height="2" bgcolor="#CCCCCC"></td>
    </tr><tr><td>&nbsp;</td>
      </tr><tr>
      <td><div align="center" class="subtitulob">Funcionalidades a serem desativadas </div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
		  <td>
		    <input type="hidden" value="<%=int_Cd_ProjetoProject%>" name="pCdProjProject">
		  
			<input type="hidden" value="<%=str_Acao%>" name="pAcao2">
			<input type="hidden" value="<%=int_Plano%>" name="pintPlano2">
			<input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject2">
			<input type="hidden" value="PDS" name="pPlano2">
			<input type="hidden" value="0" name="pCdSeqFunc2">
			
			<input type="hidden" value="<%=int_CD_Onda%>" name="pOnda">
			<input type="hidden" value="<%=str_Fase%>" name="pFase">
			<input type="hidden" value="<%=strPlanoOriginal%>" name="pPlanoOriginal">			
		  </td>
      </tr>
  </table>
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
    <tr bgcolor="#CCCCCC">
	  <%if str_Acao <> "C" then%>
      		<td width="3%" bgcolor="#9C9A9C" class="titcoltabela"><a href="inclui_altera_plano_pds_func.asp?pAcao=I&pPlano=<%=int_Plano%>&pIdTaskProject=<%=int_Id_TarefaProject%>"><img src="../img/botao_novo_off_02.gif" alt="Incluir uma nova Funcionalidade" width="34" height="23" border="0"></a></td>
      <%end if%>
	  <td width="19%">&nbsp;</td>
      <td colspan="2" class="titcoltabela"><div align="center">Desligamento</div></td>
      <td width="28%" class="titcoltabela"><div align="center">Procedimento de </div></td>
      <td width="30%" class="titcoltabela"><div align="center">Destino dos dados pr&oacute;prios </div></td>
    </tr>
    <tr bgcolor="#CCCCCC">
	  <%if str_Acao <> "C" then%>
      	<td bgcolor="#9C9A9C" class="titcoltabela">&nbsp;</td>
	  <%end if%>
      <td class="titcoltabela"><div align="center"><span class="campob">Funcionalidade</span></div></td>
      <td width="7%" class="titcoltabela"><div align="center">Data</div></td>
      <td width="9%" class="titcoltabela"> <div align="center">Hora </div></td>
      <td class="titcoltabela"><div align="center">Desligamento (descri&ccedil;&atilde;o)</div></td>
      <td class="titcoltabela"><div align="center">Tempo de reten&ccedil;&atilde;o </div></td>
    </tr>
	<%
	set rdsPds_Func = db_Cogest.Execute(str_Pds_Func)
	if not rdsPds_Func.EOF then 
	      Do while not rdsPds_Func.EOF
	%>
    <tr bgcolor="#E9E9E9">
	  <%if str_Acao <> "C" then%>
      	<td bgcolor="#9C9A9C"><a href="inclui_altera_plano_pds_func.asp?pAcao=A&pPlano=<%=int_Plano%>&pIdTaskProject=<%=int_Id_TarefaProject%>&pCdSeqFunc=<%=rdsPds_Func("PPDS_NR_SEQUENCIA_FUNC")%>"><img src="../img/botao_abrir_off_02.gif" alt="Alterar Funcionalidade" width="34" height="23" border="0"></a><a href="javascript:document.frm_Plano_PDS_Func.pCdSeqFunc2.value='<%=rdsPds_Func("PPDS_NR_SEQUENCIA_FUNC")%>';confirma_Exclusao_Parcial()"><img src="../img/botao_deletar_off_02.gif" alt="Excluir Funcionalidade" width="34" height="23" border="0"></a></td>
      <%end if%>
	  <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPds_Func("PPDS_TX_FUNC_DESATIVADAS")%></div></td>
	  <% 	strDia = ""		
	strMes = ""
	strAno = ""
	vetDtDesliga = split(Trim(rdsPds_Func("PPDS_DT_DESLIGAMENTO")),"/")						
	strDia = trim(vetDtDesliga(1))
	if cint(strDia) < 10 then
		strDia = "0" & strDia
	end if			
	strMes = trim(vetDtDesliga(0))			
	if cint(strMes) < 10 then
		strMes = "0" & strMes
	end if
	strAno = trim(vetDtDesliga(2))
	dat_DtDesliga = strDia & "/" & strMes & "/" & strAno 
   %>
      <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=dat_DtDesliga%></div></td>
      <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPds_Func("PPDS_TX_HR_DESLIGAMENTO")%></div></td>
      <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPds_Func("PPDS_TX_PROC_DESLIGAMENTO")%></div></td>
      <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPds_Func("PPDS_TX_DEST_DD_TEMPO_RETENCAO")%></div></td>
    </tr>
    <%       rdsPds_Func.movenext 
	     Loop 
	end if %>	 
  </table>
	<% end if %>
  </form>
  <% 
  	rdsPds_Func.close
	set rdsPds_Func = nothing 
	
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