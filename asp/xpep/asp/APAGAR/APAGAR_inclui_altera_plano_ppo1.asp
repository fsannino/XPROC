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
<html>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBegin template="/Templates/BASICO_XPEP_01.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
	<!-- InstanceBeginEditable name="doctitle" -->
	<title>SINERGIA # XPROC # Processos de Negócio</title>
	<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
	<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">	
	<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaSci" -->		
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
          <td width="26"><a href="javascript:confirma_ppo()"><img src="../../../imagens/continua_F02.gif" width="24" height="24" border="0"></a></td>
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
    <td><table width="75%" border="0">
        <tr>
          <td class="campo"><div align="center">A&ccedil;&atilde;o</div></td>
        </tr>
        <tr>
          <td bgcolor="#EEEEEE"> 
            <div align="center" class="campob"><%=str_Acao%></div></td>
        </tr>
      </table></td>
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
</table><hr>

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
    </tr></tr>
    <td colspan="5"><!--#include file="../includes/inc_lista_Responsavel_Sinergia_Um.asp" --></td>
    </tr>
    <tr> 
      <td colspan="5"><!--#include file="../includes/inc_lista_Responsavel_Legado.asp" --> 
    <tr> 
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right" class="campob">Tempo da Parada:</div></td>
      <td><input type="text" name="txtTempParada" size="3">
        <select name="selUnidMedida" size="1">
			<option value="Hora">Hora</option>
			<option value="Dia" selected>Dia</option>
			<option value="Mês">Mês</option>
        </select></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td valign="top">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td valign="top"><div align="right" class="campob">Procedimentos para a Parada:</div></td>
      <td><textarea name="txtProcedParada" cols="50" rows="5" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"></textarea></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="25" colspan="3" class="campo"> <table width="89%" border="0">
          <tr> 
            <td width="3%">&nbsp;</td>
            <td width="97%" class="campob">&nbsp;</td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td class="campo">&nbsp;</td>
      <td class="campo"><div align="right" class="campob">Data da Parada do  Legado:</div></td>
      <td>        <table width="100%"  border="0">
          <tr>
            <td><input type="text" name="txtDtParadaLegado" size="10" maxlength="10"></td>
            <td><div align="right"><span class="campob">Início no R/3:</span></div></td>
            <td><input type="text" name="txtDtIniR3" size="10" maxlength="10"></td>
          </tr>
        </table></td>
      <td class="campo"><div align="right" class="campob"></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td align="right" valign="top" class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div class="campob"> 
          <div align="left">Gestor do Processo:</div>
        </div></td>
      <td><!--#include file="../includes/inc_combo_Usuario_Gestor.asp" --> </td>
      <td align="right" valign="top" class="campob"><div align="right">Data Limite 
          para aprovação:</div></td>
      <td><input type="text" name="txtDtLimiteAprov" size="10" maxlength="10"></td>
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
      <td><div align="right" class="campob" >Procedimentos de Contingência:</div></td>
      <td><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></td>
      <td class="campob">Procedimento de Comunicação:</td>
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
  </table>
</form>
<!-- InstanceEndEditable -->
</body>

<!-- InstanceEnd --></html>
