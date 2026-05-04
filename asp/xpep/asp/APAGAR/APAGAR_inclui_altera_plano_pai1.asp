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
	<script language="JavaScript" src="pupdate.js"></script>
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
						else
						{
							if (strobjnome=='txtRestricoes')
							{
								document.forms[0].txtRestricoes.value = strvalor.substr(0,i);
							}
							else
							{
								document.forms[0].txtProcedimento.value = strvalor.substr(0,i);
							}
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
			var int_RespTecLegGeral = document.frm_Plano_PAI.selRespTecLegGeral.selectedIndex;	
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
					
		   //*** Responsável Legado
		   if (int_RespTecLegGeral == 0)
			  {
			  alert("É obrigatório a seleção no campo Responsável Legado!");
			  document.frm_Plano_PAI.selRespTecLegGeral.focus();
			  return;
			  }    
			  
		   //*** Responsável Sinergia
		   if (int_RespTecSinGeral == 0)
			  {
			  alert("É obrigatório a seleção no campo Responsável Sinergia!");
			  document.frm_Plano_PAI.selRespTecSinGeral.focus();
			  return;
			  }    	
			  
		   document.frm_Plano_PAI.action="grava_plano.asp?pPlano=PAI";           
		   document.frm_Plano_PAI.submit();				
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
          <td width="26"><a href="javascript:confirma_pai()"><img src="../../../imagens/continua_F02.gif" width="24" height="24" border="0"></a></td>
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
    <td><table width="75%" border="0">
        <tr>
          <td class="campo"><div align="center">A&ccedil;&atilde;o</div></td>
        </tr>
        <tr>
          <td bgcolor="#EEEEEE"> 
            <div align="center" class="campob"><%=str_Texto_Acao%></div></td>
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
    <td width="17%" bgcolor="#EEEEEE"><div align="right" class="campo">Atvidade:</div></td>
    <td colspan="3" class="campob"><%=str_NomeAtividade%></td>
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
    <td width="33%" class="campob"><%=Day(dat_Dt_Termino) & "/" & Month(dat_Dt_Termino) & "/" & Year(dat_Dt_Termino)%></td>
  </tr>
</table>
<form name="frm_Plano_PAI" method="post" action="">
  <table width="98%" border="0">      
	<tr> 
	  <td colspan="5"><hr></td>
	</tr>      
	    
    <tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Código da Interface:</div></td>
		<td><input type="text" name="txtCdInterface" value="<%=str_txtCdInterface%>"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Grupo:</div></td>
		<td><input type="text" name="txtGrupo" value="<%=str_txtGrupo%>"></td>
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
		<td><input type="text" name="txtNomeInterface" value="<%=str_txtNomeInterface%>"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Programa Envolvido:</div></td>
		<td><input type="text" name="txtPgrmEnvolv" value="<%=str_txtPgrmEnvolv%>"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob" valign="top"><div align="right">Pré-Requisitos:</div></td>
		<td><textarea type="text" cols="40" rows="4" name="txtPreRequisitos" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtPreRequisitos%></textarea></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob" valign="top"><div align="right">Restrições:</div></td>
		<td><textarea type="text" cols="40" rows="4" name="txtRestricoes" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtRestricoes%></textarea></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
        
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Dependências:</div></td>
		<td><input type="text" name="txtDependencias" value="<%=str_txtDependencias%>"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Data de Inicio:</div></td>
		<td>
			<input type="text" name="txtDtInicio_Pai" maxlength="10" size="10" readonly value="<%=str_txtDtInicio_Pai%>">
			<img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0" onClick="getCalendarFor(document.body.offsetHeight,document.frm_Plano_PAI.txtDtInicio_Pai)">
		</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob"><div align="right">Responsável pelo Acionamento:</div></td>
		<td><input type="text" name="txtRespAciona" value="<%=str_txtRespAciona%>"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
        </tr>
		<td class="campo">&nbsp;</td>
		<td class="campob" valign="top"><div align="right">Procedimento:</div></td>
		<td><textarea type="text" cols="40" rows="4" name="txtProcedimento" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"><%=str_txtProcedimento%></textarea></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
    </tr>
    
    <tr> 
      <td colspan="5">    <tr> 
      <td colspan="5"><table width="100%"  border="0">
  <tr>
    <td>&nbsp;</td>
    <td class="campob">Respons&aacute;vel Legado </td>
    <td><!--#include file="../includes/inc_lista_Responsavel_Legado_Um.asp" --></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="2%">&nbsp;</td>
    <td width="18%" class="campob">Respons&aacute;vel Sinergia </td>
    <td width="78%"><!--#include file="../includes/inc_lista_Responsavel_Sinergia_Um.asp" --></td>
    <td width="1%">&nbsp;</td>
    <td width="1%">&nbsp;</td>
  </tr>
</table>
      </td>
    </tr>
      
    <tr>
      <td>&nbsp;</td>
      <td><div align="right" class="campob" >Procedimentos de Conting&ecirc;ncia:</div></td>
      <td><img src="../../../imagens/seta.gif" width="30" height="24" border="0"></td>
      <td class="campob">Procedimento de Comunica&ccedil;&atilde;o:</td>
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
      <td>
	  	  <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject">
		  <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
		  <input type="hidden" value="<%=str_NomeAtividade%>" name="pNomeAtividade">
		  <input type="hidden" value="<%=dat_Dt_Inicio%>" name="pDtInicioAtiv">
		  <input type="hidden" value="<%=dat_Dt_Termino%>" name="pDtFimAtiv">
		  <input type="hidden" value="<%=str_Acao%>" name="pAcao">
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
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<!-- PopUp Calendar BEGIN -->
<script language="JavaScript">
if (document.all) {
 document.writeln("<div id=\"PopUpCalendar\" style=\"position:absolute; left:0px; top:0px; z-index:7; width:200px; height:77px; overflow: visible; visibility: hidden; background-color: #FFFFFF; border: 1px none #000000\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout(\'hideCalendar()\',500)\">");
 document.writeln("<div id=\"monthSelector\" style=\"position:absolute; left:0px; top:0px; z-index:9; width:181px; height:27px; overflow: visible; visibility:inherit\">");}
else if (document.layers) {
 document.writeln("<layer id=\"PopUpCalendar\" pagex=\"0\" pagey=\"0\" width=\"200\" height=\"200\" z-index=\"100\" visibility=\"hide\" bgcolor=\"#FFFFFF\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout('hideCalendar()',500)\">");
 document.writeln("<layer id=\"monthSelector\" left=\"0\" top=\"0\" width=\"181\" height=\"27\" z-index=\"9\" visibility=\"inherit\">");}
else {
 document.writeln("<p><font color=\"#FF0000\"><b>Error ! The current browser is either too old or too modern (usind DOM document structure).</b></font></p>");}
</script>
<noscript></noscript>
<table border="1" cellspacing="1" cellpadding="2" width="200" bordercolorlight="#000000" bordercolordark="#000000" vspace="0" hspace="0"><form name="ppcMonthList"><tr><td align="center" bgcolor="#CCCCCC"><a href="javascript:moveMonth('Back')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b>&lt;&nbsp;</b></font></a><font face="MS Sans Serif, sans-serif" size="1"> 
<select name="sItem" onMouseOut="if(ppcIE){window.event.cancelBubble = true;}" onChange="switchMonth(this.options[this.selectedIndex].value)" style="font-family: 'MS Sans Serif', sans-serif; font-size: 9pt"><option value="0" selected>2000
  . Janeiro</option><option value="1">2000 . Fevereiro</option><option value="2">2000
  . Março</option><option value="3">2000 . Abril</option><option value="4">2000
  . Maio</option><option value="5">2000 . Junho</option><option value="6">2000
  . Julho</option><option value="7">2000 . Agosto</option><option value="8">2000
  . Agosto</option><option value="9">2000 . Outubro</option><option value="10">2000
  . Novembro</option><option value="11">2000 . Dezembro</option><option value="0">2001
  . Janeiro</option></select></font><a href="javascript:moveMonth('Forward')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b>&nbsp;&gt;</b></font></a></td></tr></form></table>
<table border="1" cellspacing="1" cellpadding="2" bordercolorlight="#000000" bordercolordark="#000000" width="200" vspace="0" hspace="0"><tr align="center" bgcolor="#CCCCCC"><td width="20" bgcolor="#FFFFCC"><b><font face="MS Sans Serif, sans-serif" size="1">Do</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Se</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Te</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Qa</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Qi</font></b></td><td width="20"><b><font face="MS Sans Serif, sans-serif" size="1">Sx</font></b></td><td width="20" bgcolor="#FFFFCC"><b><font face="MS Sans Serif, sans-serif" size="1">Sa</font></b></td></tr></table>
<script language="JavaScript">
if (document.all) {
 document.writeln("</div>");
 document.writeln("<div id=\"monthDays\" style=\"position:absolute; left:0px; top:52px; z-index:8; width:200px; height:17px; overflow: visible; visibility:inherit; background-color: #FFFFFF; border: 1px none #000000\">&nbsp;</div></div>");}
else if (document.layers) {
 document.writeln("</layer>");
 document.writeln("<layer id=\"monthDays\" left=\"0\" top=\"52\" width=\"200\" height=\"17\" z-index=\"8\" bgcolor=\"#FFFFFF\" visibility=\"inherit\">&nbsp;</layer></layer>");}
else {/*NOP*/}
</script>
<!-- PopUp Calendar END -->
<!-- InstanceEndEditable -->
</body>

<!-- InstanceEnd --></html>
