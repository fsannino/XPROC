<%
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
<html>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBegin template="/Templates/BASICO_XPEP_01.dwt" codeOutsideHTMLIsLocked="false" -->
	<head>
		<!-- InstanceBeginEditable name="doctitle" -->
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
		<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
		<!-- InstanceEndEditable -->
	</head>
<!-- InstanceBeginEditable name="JavaSci" -->	
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
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../Funcao/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../../indexA.asp"><img src="../../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
      
    <td colspan="3" height="20"><!-- InstanceBeginEditable name="Botao_01" -->
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:confirma_pcm()"><img src="../../../imagens/continua_F02.gif" width="24" height="24" border="0"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
      <!-- InstanceEndEditable --></td>
  </tr>
</table>     
<!-- InstanceBeginEditable name="Corpo_Princ" --> 
<table width="98%" border="0" cellspacing="0" cellpadding="0">
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
</body>

<!-- InstanceEnd --></html>
