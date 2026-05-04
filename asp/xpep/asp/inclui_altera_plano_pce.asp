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
						
						if (strobjnome=='txtAtivCritica')
						{
							document.forms[0].txtAtivCritica.value = strvalor.substr(0,i);
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
			
		function confirma_pcd()
		{			
			var str_RespTecLegGeral = document.frm_Plano_PCD.txtRespTecLegGeral.value; 
			var str_RespFunLegGeral = document.frm_Plano_PCD.txtRespFunLegGeral.value; 	
			var str_RespTecSinGeral = document.frm_Plano_PCD.txtRespTecSinGeral.value; 
			var str_RespFunSinGeral = document.frm_Plano_PCD.txtRespFunSinGeral.value;	
			
			var int_DesenvAssoc		= document.frm_Plano_PCD.lstDesenvAssociadosSel.selectedIndex;						
			var str_SistLegado		= document.frm_Plano_PCD.txtSistLegado.value; 
			var str_Extracao_PCD	= document.frm_Plano_PCD.txtExtracao_PCD.value;	
			var str_Carga_PCD 		= document.frm_Plano_PCD.txtCarga_PCD.value;						
			var str_ArqCarga		= document.frm_Plano_PCD.txtArqCarga.value;				
			var str_Volume 			= document.frm_Plano_PCD.txtVolume.value;
			var str_Dependencias	= document.frm_Plano_PCD.txtDependencias.value;
			var str_DTExtracao_PCD	= document.frm_Plano_PCD.txtDTExtracao_PCD.value;
			var srt_DTCarga_PCD_Ini	= document.frm_Plano_PCD.txtDTCarga_PCD_Inicio.value;
			var srt_DTCarga_PCD_Fim	= document.frm_Plano_PCD.txtDTCarga_PCD_Fim.value;
			var str_ComoExecuta		= document.frm_Plano_PCD.txtComoExecuta.value;
						 
			//*** Responsável Técnico - Legado				  
			if(str_RespTecLegGeral == '')
			  {
			  alert("É obrigatório o preenchimento do campo Responsável Legado - Técnico!");
			  document.frm_Plano_PCD.txtRespTecLegGeral.focus();
			  return;
			  } 
			 
			//*** Responsável Funcional - Legado 
			if(str_RespFunLegGeral == '')
			  {
			  alert("É obrigatório o preenchimento do campo Responsável Legado - Funcional!");
			  document.frm_Plano_PCD.txtRespFunLegGeral.focus();
			  return;
			  } 
			
			//*** Responsável Técnico - Sinergia				  
			if(str_RespTecSinGeral == '')
			  {
			  alert("É obrigatório o preenchimento do campo Responsável Sinergia - Técnico!");
			  document.frm_Plano_PCD.txtRespTecSinGeral.focus();
			  return;
			  }
			
			//*** Responsável Funcional - Sinergia 
			if(str_RespFunSinGeral == '')
			  {
			  alert("É obrigatória a seleção de um Responsável Sinergia - Funcional!");
			  document.frm_Plano_PCD.txtRespFunSinGeral.focus();
			  return;
			  }
			 			  
			//*** Sistema Legado de Origem				  
			if(str_SistLegado == "")
			  {
			  alert("É obrigatório o preenchimento do campo Sistema Legado de Origem!");
			  document.frm_Plano_PCD.txtSistLegado.focus();
			  return;
			  } 
			  
		   //*** Extração  	
		   if (str_Extracao_PCD == "")
			  {
			  alert("É obrigatório o preenchimento do campo Extração(h)!");
			  document.frm_Plano_PCD.txtExtracao_PCD.focus();
			  return;
			  }			
		   else
			  {
				if (isNaN(str_Extracao_PCD))
				{
					alert("O contéudo do campo Extração(h) deve ser preenchido apenas com nº!");
					document.frm_Plano_PCD.txtExtracao_PCD.value = '';
					document.frm_Plano_PCD.txtExtracao_PCD.focus();
					return;
				}
			  }  
			
		   //*** Carga  	
		   if (str_Carga_PCD == "")
			  {
			  alert("É obrigatório o preenchimento do campo Carga(h)!");
			  document.frm_Plano_PCD.txtCarga_PCD.focus();
			  return;
			  }  
			else
			  {
				if (isNaN(str_Carga_PCD))
				{
					alert("O contéudo do campo Carga(h) deve ser preenchido apenas com nº!");
					document.frm_Plano_PCD.txtCarga_PCD.value = '';
					document.frm_Plano_PCD.txtCarga_PCD.focus();
					return;
				}
			  }  
			    
		   //*** Arquivos de carga   	
		   if (str_ArqCarga == "")
			  {
			  alert("É obrigatório o preenchimento do campo Arquivos de Carga!");
			  document.frm_Plano_PCD.txtArqCarga.focus();
			  return;
			  }	
		
		  //*** Volume
		   if (str_Volume == "")
			  {
			  alert("É obrigatório o preenchimento do campo Volume!");
			  document.frm_Plano_PCD.txtVolume.focus();
			  return;
			  } 
		  	else
			  {
				if (isNaN(str_Volume))
				{
					alert("O contéudo do campo Volume deve ser preenchido apenas com nº!");
					document.frm_Plano_PCD.txtVolume.value = '';
					document.frm_Plano_PCD.txtVolume.focus();
					return;
				}
			  }  
			  	
		   //*** Dependências
		   if (str_Dependencias == "")
			  {
			  alert("É obrigatório o preenchimento do campo Dependências!");
			  document.frm_Plano_PCD.txtDependencias.focus();
			  return;
			  } 	
				
		   //*** Data Extração
		   if (str_DTExtracao_PCD == "")
			  {
			  alert("É obrigatório o preenchimento do campo Data Extração!");
			  document.frm_Plano_PCD.txtDTExtracao_PCD.focus();
			  return;
			  } 	
		
		   //*** Data Carga - Inicio
		   if (srt_DTCarga_PCD_Ini == "")
			  {
			  alert("É obrigatório o preenchimento do campo Data Carga - Início!");
			  document.frm_Plano_PCD.txtDTCarga_PCD_Inicio.focus();
			  return;
			  } 	
				
		   //*** Data Carga - Fim
		   if (srt_DTCarga_PCD_Fim == "")
			  {
			  alert("É obrigatório o preenchimento do campo Data Carga - Fim!");
			  document.frm_Plano_PCD.txtDTCarga_PCD_Fim.focus();
			  return;
			  } 		
				
		   //*** Como executa
		   if (str_ComoExecuta == "")
			  {
			  alert("É obrigatório o preenchimento do campo Como Executa!");
			  document.frm_Plano_PCD.txtComoExecuta.focus();
			  return;
			  }
		   document.frm_Plano_PCE.action="grava_plano.asp?pPlano=PCE";           
		   document.frm_Plano_PCE.submit();				
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
          <td width="26"><a href="javascript:confirma_pce();"><img src="../../../imagens/continua_F02.gif" width="24" height="24" border="0"></a></td>
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
          <td width="10%" class="subtitulob">&nbsp;</td>
          <td class="subtitulob">Contingências para Estabilização - PCE</td>
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
<form name="frm1" method="post" action="">
  <table width="98%" border="0">
  <td width="2%" class="campo">&nbsp;</td>
      <td width="20%" class="campob"><div align="right">Atividade Crítica:</div></td>
      <td width="74%"><textarea type="text" cols="34" rows="4" name="txtAtivCritica" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"></textarea></td>
      <td width="2%">&nbsp;</td>
      <td width="2%">&nbsp;</td>
  </tr>
  <td class="campo">&nbsp;</td>
      <td class="campob"><div align="right">Ações Corretivas/Contingências:</div></td>
      <td><textarea type="text" cols="34" rows="4" name="txtAcoesCorrConting" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);"></textarea></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="campo">&nbsp;</td>
    <td height="25" class="campob"><div align="right">Aprovador pela Conting&ecirc;ncia:</div></td>
    <td class="campo"><!--#include file="../includes/inc_combo_Usuario_Aprov.asp" --></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="campo">&nbsp;</td>
    <td height="25" class="campob"><div align="right">Data da Aprova&ccedil;&atilde;o:</div></td>
    <td class="campo">
      <input type="text" name="txtDTAprovacao_PCE">
</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <td class="campo">&nbsp;</td>
      <td class="campob"><div align="right">Executor da Conting&ecirc;ncia :</div></td>
      <td><span class="campo">
        <!--#include file="../includes/inc_combo_Usuario_Exec.asp" -->
      </span></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
  </tr>
  <td class="campo">&nbsp;</td>
      <td class="campob"><div align="right">Controlador da Conting&ecirc;ncia:</div></td>
      <td><span class="campo">
        <!--#include file="../includes/inc_combo_Usuario_Contr.asp" -->
      </span></td>
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
  </table>
</form>
<!-- InstanceEndEditable -->
</body>

<!-- InstanceEnd --></html>
