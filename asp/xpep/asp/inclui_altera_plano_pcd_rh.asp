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

int_Cd_Projeto_Project 	= Request("pCdProjProject")
str_Atividade  			= Request("pTArefa")
str_Cd_Onda 			= Request("pOnda")
str_Cd_Plano 			= Request("pPlano")
str_Fase 				= Request("pFase")

str_Acao = Request("pAcao")
int_Cd_ProjetoProject = Request("pCdProjProject")
int_CD_Onda = request("pOnda")

if request("pPlano") <> "" then
	int_Plano = request("pPlano")
else
	int_Plano = request("pintPlano")
end if

if Request("pTArefa") <> "" then
	int_Id_TarefaProject = Request("pTArefa")
else
	int_Id_TarefaProject = Request("idTaskProject")
end if

if str_Acao = "I" then
   str_Texto_Acao = "Inclusão"
else
    str_Texto_Acao = "Alteração"
end if

'*********** SELECIONA SIGLA E DESCRIÇÃO DO PLANO SELECIONADO OARA A CRIAÇÃO DO PCM ****
str_SelPlano = ""
str_SelPlano = str_SelPlano & "SELECT PLAN_TX_SIGLA_PLANO, PLAN_TX_DESCRICAO_PLANO " 
str_SelPlano = str_SelPlano & " FROM XPEP_PLANO_ENT_PRODUCAO "
str_SelPlano = str_SelPlano & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & int_Plano

'Response.write str_SelPlano & "<br>"

set rds_SelPlano = db_Cogest.Execute(str_SelPlano)
if not rds_SelPlano.Eof then
   str_Plano = rds_SelPlano("PLAN_TX_SIGLA_PLANO") & " - " &  rds_SelPlano("PLAN_TX_DESCRICAO_PLANO")
else
   str_Plano = ""   
end if

'Response.write str_Plano
'Response.end

rds_SelPlano.close
set rds_SelPlano = nothing

sql_PCD = ""
sql_PCD = sql_PCD & "SELECT PLAN_NR_SEQUENCIA_PLANO"			
sql_PCD = sql_PCD & ", PLTA_NR_SEQUENCIA_TAREFA "	
sql_PCD = sql_PCD & ", PPCD_NR_SEQUENCIA_FUNC "	
sql_PCD = sql_PCD & ", PPCD_TX_SISTEMA_LEGADO "			
sql_PCD = sql_PCD & ", PPCD_TX_DADO_A_SER_MIGRADO "
sql_PCD = sql_PCD & ", PPCD_TX_TIPO_ATIVIDADE "
sql_PCD = sql_PCD & ", PPCD_TX_TIPO_DADO "
sql_PCD = sql_PCD & ", PPCD_TX_CARAC_DADO "
sql_PCD = sql_PCD & ", PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO "
sql_PCD = sql_PCD & ", PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO "
sql_PCD = sql_PCD & ", PPCD_TX_QTD_TEMPO_EXEC_CARGA "
sql_PCD = sql_PCD & ", PPCD_TX_UNID_TEMPO_EXEC_CARGA "			
sql_PCD = sql_PCD & ", PPCD_TX_ARQ_CARGA "
sql_PCD = sql_PCD & ", PPCD_NR_VOLUME "
sql_PCD = sql_PCD & ", PPCD_TX_DEPENDENCIAS "
sql_PCD = sql_PCD & ", PPCD_DT_EXTRACAO "
sql_PCD = sql_PCD & ", PPCD_DT_CARGA_INICIO "
sql_PCD = sql_PCD & ", PPCD_DT_CARGA_FIM "
sql_PCD = sql_PCD & ", PPCD_TX_COMO_EXECUTA "
sql_PCD = sql_PCD & ", PPCD_NR_ID_PLANO_CONTINGENCIA "
sql_PCD = sql_PCD & ", PPCD_NR_ID_PLANO_COMUNICACAO "				
sql_PCD = sql_PCD & ", USUA_CD_USUARIO_RESP_LEG_TEC "
sql_PCD = sql_PCD & ", USUA_CD_USUARIO_RESP_LEG_FUN "			
sql_PCD = sql_PCD & ", USUA_CD_USUARIO_RESP_SIN_TEC "	
sql_PCD = sql_PCD & ", USUA_CD_USUARIO_RESP_SIN_FUN "		
sql_PCD = sql_PCD & ", ATUA_TX_OPERACAO "
sql_PCD = sql_PCD & ", ATUA_CD_NR_USUARIO "
sql_PCD = sql_PCD & ", ATUA_DT_ATUALIZACAO "
sql_PCD = sql_PCD & " FROM XPEP_PLANO_TAREFA_PCD_FUNC"
sql_PCD = sql_PCD & " WHERE PLTA_NR_SEQUENCIA_TAREFA = " & int_Id_TarefaProject

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

'RESPONSE.Write(str_Sql_DadosAdicionais_Tarefa)
'response.End()
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
rds_Responsavel.Close
set rds_Responsavel = Nothing
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
<script language="JavaScript">	
	function Habilita(form)
	{
	if ( form.tipo.value == 2)
		{
		form.cdassi.disabled = false
		form.cdassi.style.backgroundColor = "#FFFFFF"
		}
	else
		{
		form.cdassi.disabled = true
		form.cdassi.style.backgroundColor = "#CCCCCC"
		}
	}
	
	function confirma_Exclusao_Ativ()
	{
		  if(confirm("Confirma a exclusão deste Registro?"))
		  {
			document.frm_Plano_PCD.pAcao.value = 'E';		
			document.frm_Plano_PCD.pOndaRH.value = 'RH';	
			document.frm_Plano_PCD.action='grava_plano.asp?pPlano=PCD' 			        
			document.frm_Plano_PCD.submit();
		  }
	}
	
	function confirma_Exclusao_Sub(intPlano,int_CD_Onda,intDesenv)
	{
		//alert(intPlano + ' - ' + int_CD_Onda + ' - ' + intDesenv)	
		  if(confirm("Confirma a exclusão deste Registro?"))
		  {	  
		    document.frm_Plano_PCD.pAcao.value = 'E';
			document.frm_Plano_PCD.pintPlano.value = intPlano;			
			document.frm_Plano_PCD.pOnda.value = int_CD_Onda;	
			document.frm_Plano_PCD.pDesenv.value = intDesenv;		
			document.frm_Plano_PCD.action='grava_sub_ativ.asp' 			        
			document.frm_Plano_PCD.submit();	
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
<form name="frm_Plano_PCD">
    <table width="88%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="6%">&nbsp;</td>
        <td width="81%" class="subtitulo"><table width="75%" border="0" cellpadding="0" cellspacing="7">
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
        <td width="13%"><table width="96%"  border="0">
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
        <td width="17%" bgcolor="#EEEEEE">
          <div align="right" class="campo">Atividade:</div></td>
        <td colspan="3" class="campob"><%=str_NomeAtividade%></td>
      </tr>
      <tr>
        <td bgcolor="#EEEEEE"><div align="right" class="campo">Respons&aacute;vel:</div></td>
        <td colspan="3" class="campob"><%=str_Nome_Responsavel%></td>
      </tr>
      <%   
  Function FormataData(str_Data)	
	 if cint(Day(str_Data)) < 10 then 
		strDiaInicio = "0" & Day(str_Data)
	 else
		strDiaInicio = Day(str_Data)
	 end if
	  
	 if cint(Month(str_Data)) < 10 then 
		strMesInicio = "0" & Month(str_Data)
	 else
		strMesInicio = Month(str_Data)
	 end if
	 FormataData = strDiaInicio & "/" & strMesInicio & "/" & Year(str_Data)	
  end function
  
  
  '*** DATA IN&Iacute;CIO
  'if cint(Day(dat_Dt_Inicio)) < 10 then 
  '	strDiaInicio = "0" & Day(dat_Dt_Inicio)
  'else
  '	strDiaInicio = Day(dat_Dt_Inicio)
  'end if
  
  'if cint(Month(dat_Dt_Inicio)) < 10 then 
  '	strMesInicio = "0" & Month(dat_Dt_Inicio)
  'else
  '	strMesInicio = Month(dat_Dt_Inicio)
  'end if
  if dat_Dt_Inicio <> "" then
  	dat_Dt_Inicio = FormataData(dat_Dt_Inicio) 'strDiaInicio & "/" & strMesInicio & "/" & Year(dat_Dt_Inicio)
  end if
  
  '*** DATA FIM
  
  if dat_Dt_Inicio <> "" then
  	dat_Dt_Termino = FormataData(dat_Dt_Termino) 'strDiaFim & "/" & strMesFim & "/" & Year(dat_Dt_Termino)
  end if
  
  'if cint(Day(dat_Dt_Termino)) < 10 then 
  '	strDiaFim = "0" & Day(dat_Dt_Termino)
  'else
  '	strDiaFim = Day(dat_Dt_Termino)
  'end if
  
  'if cint(Month(dat_Dt_Termino)) < 10 then 
  '	strMesFim = "0" & Month(dat_Dt_Termino)
  'else
  '	strMesFim = Month(dat_Dt_Termino)
  'end if
  'dat_Dt_Termino = strDiaFim & "/" & strMesFim & "/" & Year(dat_Dt_Termino)
  %>
      <tr>
        <td bgcolor="#EEEEEE">
          <div align="right" class="campo">Data In&iacute;cio:</div></td>
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
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="2" bgcolor="#CCCCCC"></td>
      </tr>
      <tr>
        <td>
		 <input type="hidden" value="<%=int_Cd_ProjetoProject%>" name="pCdProjProject">
	          <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
			  <input type="hidden" value="<%=str_Acao%>" name="pAcao">
              <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
              <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject2">
              <input type="hidden" value="PCD" name="pPlano">
			  <input type="hidden" value="" name="pOnda"> 
			  <input type="hidden" value="" name="pDesenv">	
			  <input type="hidden" value="<%=str_Fase%>" name="pFase">
			  <input type="hidden" value="" name="pOndaRH">					
			  </td>   
      </tr>
    </table>
   
    <table width="100%"  border="0" cellpadding="0" cellspacing="0">
      <%if str_Acao = "A" then%>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td colspan="3" align="left" valign="bottom">
          <div class="campob">
            <table width="53%" border="0">
              <tr>
                <td width="87%">Link com Plano de A&ccedil;&otilde;es Corretivas / Conting&ecirc;ncia (PAC):</td>
                <td width="13%"><div class="campob"><a href="encaminha_plano.asp?selTipoCadastro=PAC&pSiglaPlano=PAC&pAtividade_Origen=<%="PCD - " & str_NomeAtividade%>&selOnda=<%=Trim(int_CD_Onda)%>&selPlano=<%=int_Plano%>|0&selTask1=<%=int_Id_TarefaProject%>|<%=int_ResData%>"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></div></td>
              </tr>
            </table>
        </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td width="22">&nbsp;</td>
      </tr>
      <tr>
        <td width="17">&nbsp;</td>
        <td width="22">&nbsp;</td>
        <td width="188">
          <%if str_Acao = "A" and str_Acao <> "C" then%>
          <a href="javascript:confirma_Exclusao_Ativ();"><img src="../img/botao_excluir.gif" width="85" height="19" border="0"></a>
          <%end if%>
        </td>
        <td width="146"><%if str_Acao = "C" then%>
            <a href="javascript:history.go(-1)"><img src="../img/botao_voltar.gif" width="85" height="19" border="0"></a>
            <%end if%>
        </td>
      </tr>
      <%end if%>
      <tr>
        <td width="22">&nbsp;</td>
      </tr>
    </table>
    <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">	
      <tr bgcolor="#CCCCCC"> 
        <td width="8%" bgcolor="#9C9A9C" class="titcoltabela">
		<a href="inclui_altera_plano_pcd_sub.asp?pAcao=I&pCdProjProject=<%=int_Cd_Projeto_Project%>&pTArefa=<%=str_Atividade%>&pOnda=<%=str_Cd_Onda%>&pPlano=<%=str_Cd_Plano%>&pFase=<%=str_Fase%>"><img src="../img/botao_novo_off_02.gif" alt="Incluir uma nova Atividade para o PCD" width="34" height="23" border="0"></a></td>
		<td colspan="4"><div align="center"></div>          
        <div align="center" class="campob">&nbsp;</div>          
        <div align="center"></div></td>
      </tr>
      <tr bgcolor="#CCCCCC">
        <td bgcolor="#9C9A9C" class="titcoltabela">&nbsp;</td>
        <td width="23%" class="titcoltabela"><div align="center"><span class="campob">Dado a ser Migrado</span></div></td>
        <td width="26%" class="titcoltabela"><div align="center">Sistemas Legados de Origem </div></td>
        <td width="24%" class="titcoltabela"><div align="center">Tipo de Atividade para a Carga</div></td>
        <td width="19%" class="titcoltabela"><div align="center">Tipo de Dado</div></td>
      </tr>
      <%
	'Response.write sql_PCD 
	'Response.end
	  
	set rdsPCD = db_Cogest.Execute(sql_PCD)
	if not rdsPCD.EOF then 
	      Do while not rdsPCD.EOF
	%>
      <tr bgcolor="#E9E9E9">
        <td bgcolor="#9C9A9C">
		<a href="inclui_altera_plano_pcd_sub.asp?pAcao=A&pCdProjProject=<%=int_Cd_Projeto_Project%>&pTArefa=<%=str_Atividade%>&pOnda=<%=str_Cd_Onda%>&pPlano=<%=str_Cd_Plano%>&pFase=<%=str_Fase%>&pCdSeqPCD=<%=rdsPCD("PPCD_NR_SEQUENCIA_FUNC")%>"><img src="../img/botao_abrir_off_02.gif" alt="Alterar Atividade do PCD" width="34" height="23" border="0"></a>
		<a href="javascript:confirma_Exclusao_Sub('<%=int_Plano%>','<%=str_Cd_Onda%>','<%=rdsPCD("PPCD_NR_SEQUENCIA_FUNC")%>');"><img src="../img/botao_deletar_off_02.gif" alt="Excluir Atividade do PCD" width="34" height="23" border="0"></a></td>
		<td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPCD("PPCD_TX_DADO_A_SER_MIGRADO")%></div></td>
        <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPCD("PPCD_TX_SISTEMA_LEGADO")%></div></td>
        <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPCD("PPCD_TX_TIPO_ATIVIDADE")%></div></td>
        <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPCD("PPCD_TX_TIPO_DADO")%></div></td>
      </tr>
      <%       rdsPCD.movenext 
	     Loop 
	end if	
	rdsPCD.close
	set rdsPCD = nothing
	%>
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
