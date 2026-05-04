<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

set db_Cronograma = Server.CreateObject("ADODB.Connection")
db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

str_Acao = Request("pAcao")
int_Cd_ProjetoProject = Request("pCdProjProject")
int_Id_TarefaProject = Request("pTArefa")
int_CD_Onda = request("pOnda")
intResData = request("pResData")
int_Plano = request("pPlano")

if str_Acao = "I" then
   str_Acao = "Inclusão"
else
   str_Acao = "Alteração"
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
<script language="JavaScript">

	function confirma_pcm()
	{			
		//var str_Comunicacao 	= document.frm_Plano_PCM.selComunicacao.selectedIndex; 
		var str_OqueComunicar 	= document.frm_Plano_PCM.txtOqueComunicar.value; 					
		var str_AQuemComunicar 	= document.frm_Plano_PCM.txtAQuemComunicar.value; 
		var str_UnidadeOrgao	= document.frm_Plano_PCM.txtUnidadeOrgao.value;		
		var str_QuandoOcorre  	= document.frm_Plano_PCM.txtQuandoOcorre.value;
						
		var str_RespConteudo	= document.frm_Plano_PCM.txtRespConteudo.value;			
		var str_RespDivulg 		= document.frm_Plano_PCM.txtRespDivulg.value;					
		var str_AprovadorPB		= document.frm_Plano_PCM.txtAprovadorPB.value;
		var str_txtDtAprovacao	= document.frm_Plano_PCM.txtDtAprovacao.value;
		
		//*** Comunicação				  
		/*
		if(str_Comunicacao == 0)
		  {
		  alert("É obrigatória a seleção de uma Comunicação!");
		  document.frm_Plano_PCM.selComunicacao.focus();
		  return;
		  }
		 */
		  
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
		
	   //*** DQuando Ocorre
	   if (str_QuandoOcorre == "")
		  {
		  alert("É obrigatório o preenchimento do Quando Ocorre!");
		  document.frm_Plano_PCM.txtQuandoOcorre.focus();
		  return;
		  } 
	
	  //*** Responsável pelo Conteúdo
	   if (str_RespConteudo == "")
		  {
		  alert("É obrigatório o preenchimento do campo Responsável pelo Conteúdo!");
		  document.frm_Plano_PCM.txtRespConteudo.focus();
		  return;
		  } 
	
	   //*** Responsável pela divulgação	
	   if (str_RespDivulg == "")
		  {
		  alert("É obrigatório o preenchimento do campo Responsável pela Divulgação!");
		  document.frm_Plano_PCM.txtRespDivulg.focus();
		  return;
		  } 		
			
	   //*** Aprovador PB
	   if (str_AprovadorPB == "")
		  {
		  alert("É obrigatório o preenchimento do campo Aprovador PB!");
		  document.frm_Plano_PCM.txtAprovadorPB.focus();
		  return;
		  }   
		
		//*** Data Aprovação
	   if (str_txtDtAprovacao == "")
		  {
		  alert("É obrigatório o preenchimento do campo Data de Aprovação!");
		  document.frm_Plano_PCM.txtDtAprovacao.focus();
		  return;
		  }   		
		
	   document.frm_Plano_PCM.action="grava_plano.asp?pPlano=PCM";           
	   document.frm_Plano_PCM.submit();				
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
	<!-- InstanceBeginEditable name="corpo" --><table width="625" border="0" align="center">
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
</table><table width="98%" border="0" cellspacing="0" cellpadding="0">
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
          <td class="subtitulo">Plano de Comunicação - PCM</td>
        </tr>
      </table></td>
    <td width="13%"><table width="75%" border="0">
        <tr> 
          <td class="campo"><div align="center">A&ccedil;&atilde;o</div></td>
        </tr>
        <tr> 
          <td bgcolor="#EEEEEE"> <div align="center" class="campob"><%=str_Acao%></div></td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="75%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <tr> 
    <td width="17%" bgcolor="#EEEEEE"><div align="right" class="campo">Atividade:</div></td>
    <td colspan="3" class="subtitulob"><%=str_NomeAtividade%></td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Respons&aacute;vel:</div></td>
    <td colspan="3" class="campob"><%=str_Nome_Responsavel%></td>
  </tr>
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Data In&iacute;cio:</div></td>
    <td width="21%" class="campob"><%=Day(dat_Dt_Inicio) & "/" & Month(dat_Dt_Inicio) & "/" & Year(dat_Dt_Inicio)%></td>
    <td width="20%" bgcolor="#EEEEEE"> <div align="right" class="campo">Data de 
        T&eacute;rmino:</div></td>
    <td width="33%" class="campob"><%=Day(dat_Dt_Termino) & "/" & Month(dat_Dt_Termino) & "/" & Year(dat_Dt_Termino) %></td>
  </tr>  
</table>
<hr>
<form name="frm_Plano_PCM" method="post" action="">
  <table width="90%" border="0" cellspacing="5" cellpadding="2">
    <tr> 
      <td colspan="3"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Equipe 
        respons&aacute;vel pela Implementa&ccedil;&atilde;o </strong></font></td>
      <td width="3%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="19%">&nbsp;</td>
      <td width="71%">&nbsp;</td>
      <td width="7%">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Comunicação:</div></td>
      <td><select name="selComunicacao" class="cmdOnda">
          <option value="INT" selected>Interna</option>
          <option value="EXT">Externa</option>
        </select></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td valign="top" class="campob"> <div align="right">O que comunicar:</div></td>
      <td>        <table width="94%"  border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="66%"><textarea cols="40" rows="4" name="txtOqueComunicar"></textarea></td>
            <td width="34%">
              <table width="100%"  border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td class="campo"><div align="center">Anexa arquivos </div></td>
                </tr>
                <tr>
                  <td><div align="center"><img src="../img/anexa_arquivo_01.gif" width="32" height="27"></div></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Para quem comunicar:</div></td>
      <td><input type="text" name="txtAQuemComunicar"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Unidade &Oacute;rg&atilde;o:</div></td>
      <td><input type="text" name="txtUnidadeOrgao"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Quando ocorre:</div></td>
      <td><input type="text" name="txtQuandoOcorre"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Frente 
        de comunica&ccedil;&atilde;o</strong></font></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Respons&aacute;vel pelo Conteúdo:</div></td>
      <td><input type="text" name="txtRespConteudo" value=""> </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Respons&aacute;vel pela Divulgação:</div></td>
      <td><input type="text" name="txtRespDivulg" value=""></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Aprovador PB:</div></td>
      <td><input type="text" name="txtAprovadorPB" value=""></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campob"><div align="right">Data de Aprovação:</div></td>
      <td><input type="text" name="txtDtAprovacao" size="10" maxlength="10"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
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
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
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
