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
	<script src="../js/troca_lista.js" language="javascript"></script>
	<script src="../js/global.js" language="javascript"></script>	
	<script language="JavaScript">
	function confirma_pcd()
	{				
		/*
		var str_CdInterface 	= document.frm_Plano_PAI.txtCdInterface.value; 					
		var str_Grupo 			= document.frm_Plano_PAI.txtGrupo.value; 		
		var str_TipoBatch		= document.frm_Plano_PAI.selTipoBatch.selectedIndex;			
		var str_NomeInterface  	= document.frm_Plano_PAI.txtNomeInterface.value;						
		var str_PgrmEnvolv		= document.frm_Plano_PAI.txtPgrmEnvolv.value;			
		var str_PreRequisitos	= document.frm_Plano_PAI.txtPreRequisitos.value;		
		var str_Restricoes		= document.frm_Plano_PAI.txtRestricoes.value;		
		var str_Dependencias	= document.frm_Plano_PAI.txtDependencias.value;
		var str_DtInicio_Pai	= document.frm_Plano_PAI.txtDtInicio_Pai.value;		
		var str_RespAciona		= document.frm_Plano_PAI.txtRespAciona.value;
		var str_Procedimento	= document.frm_Plano_PAI.txtProcedimento.value;
		var int_RespTecSinGeral	= document.frm_Plano_PAI.selRespTecSinGeral.selectedIndex;	
		var int_RespTecSinGeral	= document.frm_Plano_PAI.selRespTecSinGeral.selectedIndex;	
							  
		//*** Código da Interface
		if (str_CdInterface == "")
		  {
		  alert("É obrigatório o preenchimento do campo Código da Interface!");
		  document.frm_Plano_PAI.txtCdInterface.focus();
		  return;
		  }		
								 
	   //*** Grupo
	   if (str_Grupo == "")
		  {
		  alert("É obrigatório o preenchimento do campo Grupo!");
		  document.frm_Plano_PAI.txtGrupo.focus();
		  return;
		  }
				 
	   //*** Tipo 	
	   if (str_TipoBatch == "")
		  {
		  alert("É obrigatório o preenchimento do campo Tipo!");
		  document.frm_Plano_PAI.str_TipoBatch.focus();
		  return;
		  }  								
		
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
		  
	   //*** Data Inicio
	   if (str_DtInicio_Pai == "")
		  {
		  alert("É obrigatório o preenchimento do campo Data Início!");
		  document.frm_Plano_PAI.txtDtInicio_Pai.focus();
		  return;
		  }  
		  
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
		  
	   document.frm_Plano_PAI.action="grava_plano.asp?pPlano=PCM";           
	   document.frm_Plano_PAI.submit();				
	*/
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
          <td width="26"><a href="javascript:confirma_pcd()"><img src="../../../imagens/continua_F02.gif" width="24" height="24" border="0"></a></td>
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
          <td class="subtitulo">Plano de Convers&otilde;es de Dados - PCD</td>
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
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo">&nbsp;</td>
    <td>&nbsp;</td>
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
  <tr> 
    <td bgcolor="#EEEEEE"> <div align="right" class="campo">Data In&iacute;cio:</div></td>
    <td width="21%" class="campob"><%=Day(dat_Dt_Inicio) & "/" & Month(dat_Dt_Inicio) & "/" & Year(dat_Dt_Inicio)%></td>
    <td width="20%" bgcolor="#EEEEEE"><div align="right" class="campo">Data de T&eacute;rmino:</div></td>
    <td width="33%" class="campob"><%=Day(dat_Dt_Termino) & "/" & Month(dat_Dt_Termino) & "/" & Year(dat_Dt_Termino)%></td>
  </tr>
</table>
<form name="frm1" method="post" action="">
  <table width="98%" border="0">
      
	<tr> 
	  <td colspan="5"><hr></td>
	</tr>      
	        
    <tr> 
      <td colspan="5"><!--#include file="../includes/inc_lista_Responsavel_Legado.asp" -->
    <tr> 
      <td colspan="5"><!--#include file="../includes/inc_lista_Responsavel_Sinergia.asp" -->
<!--#include file="../includes/inc_lista_desenvolvimentos.asp" --></td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td height="25" class="campob" align="right">Dado a ser Migrado: </td>
      <td><input name="txtDadoMigrado" type="text"></td>
      <td class="campo">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td height="25" class="campob">&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campo">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td height="25" class="campob"><div align="right">Sistema Legados de Origem:</div></td>
      <td>
        <select name="selSistLegado" class="cmdOnda">
          <option value="1">== Selecione um Sistema ==</option>
          <%
			rcs_SistLegado.MoveFirst
			do while not rcs_SistLegado.eof%>
          <option value="<%=rcs_SistLegado("SIST_NR_SEQUENCIAL_SISTEMA_LEGADO")%>"><%=rcs_SistLegado("SIST_TX_CD_SISTEMA") & " - " & rcs_SistLegado("SIST_TX_DESC_SISTEMA_LEGADO")%></option>
          <%
				rcs_SistLegado.MoveNext
			loop
			 		 
			rcs_SistLegado.close()
			set rcs_SistLegado = nothing
			%>
        </select>
      </td>
      <td class="campob"><div align="right">Tipo de Ativ.para carga:</div></td>
      <td>
        <select name="selTipoCarga" class="cmdOnda">
          <option value="Manual" selected>Manual</option>
          <option value="Autom&aacute;tica">Autom&aacute;tica</option>
          <option value="Customizada">Customizada</option>
          <option value="Verifica&ccedil;&atilde;o">Verifica&ccedil;&atilde;o</option>
        </select>
      </td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td height="25" class="campo">&nbsp;</td>
      <td class="campo">&nbsp;</td>
      <td class="campo">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td height="25" class="campob"><div align="right">Tipo de dado:</div></td>
      <td class="campo">
        <select name="selTipoDados">
          <option value="Texto" selected>Texto</option>
          <option value="Idoc">Idoc</option>
        </select>
      </td>
      <td class="campob"><div align="right">Caracter&iacute;stica do dado:</div></td>
      <td>
        <select name="selCaractDado">
          <option value="Mestre" selected>Mestre</option>
          <option value="Transacional">Transacional</option>
        </select>
      </td>
    </tr>
    <tr>
      <td height="25" colspan="3" class="campo">
        <table width="89%" border="0">
          <tr>
            <td>&nbsp;</td>
            <td class="campob">&nbsp;</td>
          </tr>
          <tr>
            <td width="15%">&nbsp;</td>
            <td width="85%" class="campob">Tempo de Execu&ccedil;&atilde;o</td>
          </tr>
      </table></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td class="campo">&nbsp;</td>
      <td class="campo">
        <div align="right">Extra&ccedil;&atilde;o (h):</div></td>
      <td><input type="text" name="txtExtracao_PCD"></td>
      <td class="campo"><div align="right">Carga (h):</div></td>
      <td><input type="text" name="txtCarga_PCD"></td>
    </tr>
      
    <tr> 
      <td colspan="5"><hr></td>
    </tr>      
        
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%"><div align="right" class="campob">Arquivos de Carga:</div></td>
      <td width="18%"><input type="text" name="txtArqCarga" maxlength="30"></td>
      <td width="19%"><div class="campob" align="right">Volume:</div></td>
      <td width="44%"><input type="text" size="15" name="txtVolume" maxlength="10"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right" class="campob">Dependências:</div></td>
      <td><textarea name="txtDependencias" cols="40" rows="4"></textarea></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
      
    <tr> 
      <td height="25" colspan="3" class="campo"> 
		<table width="89%" border="0">
          <tr> 
            <td width="15%">&nbsp;</td>
            <td width="85%" class="campob">Datas:</td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
        
    <tr>
	  <td>&nbsp;</td>
	  <td><div align="right" class="campob">Extração:</div></td>
      <td><input type="text" name="txtDTExtracao_PCD"></td>
      <td><div align="right" class="campob">Carga:</div></td>
      <td><input type="text" name="txtDTCarga_PCD"></td>
    </tr>
    
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right" class="campob">Como Executa:</div></td>
      <td><textarea name="txtComoExecuta" cols="40" rows="4"></textarea></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    
    <tr> 	  
      <td colspan="5">&nbsp;</td>
    </tr>
               
    <tr> 
      <td>&nbsp;</td>
      <td><div class="campob" align="right">Ação Corretiva:</div></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    
    <tr> 
		<td>&nbsp;</td>	  
		<td colspan="2"><div align="right" class="campob" >Procedimentos de Contingência:</div></td>
	    <td><a href="inclui_plano_pce.asp"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
    </tr>  
    <tr>
		 <td>&nbsp;</td>
		 <td colspan="2"><div align="right" class="campob">Procedimento de Comunicação:</div></td>
	     <td><a href="inclui_plano_pcm.asp"><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></a></td>
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
      <td>  
	  	<input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
		  <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
		  <input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
		  <input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
		  <input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">
	  </td>
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
