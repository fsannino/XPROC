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

str_Acao = Request("pAcao") ' I incluir A alterar
int_Plano = Request("pPlano") ' codigo do palno
int_IdTaskProject = Request("pIdTaskProject") ' id da tarefa do Project
int_CdSeqFunc = Request("pCdSeqFunc") ' somente virá quando for alteração

str_FuncDesat = ""
dat_DtDesliga = ""
hor_HrDesliga = ""
hor_MnDesliga = ""
str_ProcDesl = ""
str_DestDados = ""

if str_Acao = "I" then
   	str_Texto_Acao = "Inclusão"   
else
   	str_Texto_Acao = "Alteração"   	
  	str_SQL = ""
	str_SQL = str_SQL & " Select "
	str_SQL = str_SQL & " PPDS_TX_FUNC_DESATIVADAS"
	str_SQL = str_SQL & " ,PPDS_DT_DESLIGAMENTO"
	str_SQL = str_SQL & " ,PPDS_TX_HR_DESLIGAMENTO"
	str_SQL = str_SQL & " ,PPDS_TX_PROC_DESLIGAMENTO"
	str_SQL = str_SQL & " ,PPDS_TX_DEST_DD_TEMPO_RETENCAO"
	str_SQL = str_SQL & " FROM XPEP_PLANO_TAREFA_PDS_FUNC"
	str_SQL = str_SQL & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & int_Plano
	str_SQL = str_SQL & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_IdTaskProject
	str_SQL = str_SQL & " AND PPDS_NR_SEQUENCIA_FUNC = " & int_CdSeqFunc
	'RESPONSE.Write(str_SQL)
	'RESPONSE.End()
	Set rdsPDS = db_Cogest.Execute(str_SQL)	
	str_FuncDesat = rdsPDS("PPDS_TX_FUNC_DESATIVADAS")
	dat_DtDesliga = rdsPDS("PPDS_DT_DESLIGAMENTO")
	hor_HrDesliga = Left(rdsPDS("PPDS_TX_HR_DESLIGAMENTO"),2)
	hor_MnDesliga = Right(rdsPDS("PPDS_TX_HR_DESLIGAMENTO"),2)
	str_ProcDesl = rdsPDS("PPDS_TX_PROC_DESLIGAMENTO")
	str_DestDados = rdsPDS("PPDS_TX_DEST_DD_TEMPO_RETENCAO")
	
	strDia = ""		
	strMes = ""
	strAno = ""
	vetDtDesliga = split(Trim(dat_DtDesliga),"/")							
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
end if

rdsPDS.close
set rdsPDS= nothing
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
	<script language="javascript" src="../js/digite-cal.js"></script>
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
					
					if (strobjnome=='txtProcDesl')
					{
						document.forms[0].txtProcDesl.value = strvalor.substr(0,i);
					}
					else
					{
						document.forms[0].txtDestDados.value = strvalor.substr(0,i);
					}
					break;
				}
			}
		}		
	}
		
	function confirma_pds_sub()
	{		

       if (document.frm_Cad_PDS_Func.txtFuncDesat.value == "")
	      {
		  alert("É obrigatório o preenchimento do campo Funcionalidade a ser desativada!");
		  document.frm_Cad_PDS_Func.txtFuncDesat.focus();
		  return;
		  }

       if (document.frm_Cad_PDS_Func.txtDtDesliga.value == "")
	      {
		  alert("É obrigatório o preenchimento do campo Data desativação!");
		  document.frm_Cad_PDS_Func.txtDtDesliga.focus();
		  return;
		  }

       if (document.frm_Cad_PDS_Func.txtHrDesliga.value == "")
	      {
		  alert("É obrigatório o preenchimento do campo Hora desativação!");
		  document.frm_Cad_PDS_Func.txtHrDesliga.focus();
		  return;
		  }

       if (document.frm_Cad_PDS_Func.txtHrDesliga.value > 23)
	      {
		  alert("Este campo Hora deverá ser preenchido no intervado de 0 a 23!");
		  document.frm_Cad_PDS_Func.txtHrDesliga.focus();
		  return;
		  }

       if (document.frm_Cad_PDS_Func.txtMnDesliga.value == "")
	      {
		  alert("É obrigatório o preenchimento do campo Minuto desativação!");
		  document.frm_Cad_PDS_Func.txtMnDesliga.focus();
		  return;
		  }

       if (document.frm_Cad_PDS_Func.txtMnDesliga.value > 59)
	      {
		  alert("Este campo Minuto deverá ser preenchido no intervado de 0 a 59!");
		  document.frm_Cad_PDS_Func.txtMnDesliga.focus();
		  return;
		  }

	   if (document.frm_Cad_PDS_Func.txtProcDesl.value == "")
	      {
		  alert("É obrigatório o preenchimento do campo Procedimento de desligamento");
		  document.frm_Cad_PDS_Func.txtProcDesl.focus();
		  return;
		  }

	   if (document.frm_Cad_PDS_Func.txtDestDados.value == "")
	      {
		  alert("É obrigatório o preenchimento do campo Destino dos dados");
		  document.frm_Cad_PDS_Func.txtDestDados.focus();
		  return;
		  }
	
	   document.frm_Cad_PDS_Func.action="grava_sub_ativ.asp?pPlano=PDS";           
	   document.frm_Cad_PDS_Func.submit();				
	}	

	function Verifica_Dif_Numeros(strValor,strNome,obj)	
	{		
	var ppcSV=null;
	ppcSV = obj
		if (isNaN(strValor))
		{
			alert("O contéudo deste campo deve ser preenchido apenas com números!");			
			ppcSV.value = '';
			ppcSV.focus();
			return;
		}
	}	
	
	function pega_tamanho(strCampo)
	{	
		if (strCampo == 'txtProcDesl')
		{
			valor = document.forms[0].txtProcDesl.value.length;
			document.forms[0].txttamanhoProcDesl.value = valor;
			if (valor > 300)
			{
				str1 = document.forms[0].txtProcDesl.value;
				str2 = str1.slice(0,300);
				document.forms[0].txtProcDesl.value = str2;
				valor = str2.length;
				document.forms[0].txttamanhoProcDesl.value = valor;
			}
		}
		
		if (strCampo == 'txtDestDados')
		{
			valor = document.forms[0].txtDestDados.value.length;
			document.forms[0].txttamanhoDestDados.value = valor;
			if (valor > 300)
			{
				str1 = document.forms[0].txtDestDados.value;
				str2 = str1.slice(0,300);
				document.forms[0].txtDestDados.value = str2;
				valor = str2.length;
				document.forms[0].txttamanhoDestDados.value = valor;
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
    <form name="frm_Cad_PDS_Func" method="post" action="">	
    <table width="100%"  border="0" cellspacing="0" cellpadding="5">
      <tr>
        <td>&nbsp;</td>
        <td class="subtitulob"><span class="subtitulo">Plano de Desligamento de Sistemas Legados - PDS</span></td>
        <td><table width="100%"  border="0">
          <tr>
            <td><div align="center" class="campo">A&ccedil;&atilde;o</div></td>
          </tr>
          <tr>
            <td width="103" bgcolor="#EFEFEF"><div align="center"><span class="campob"><%=str_Texto_Acao%></span></div></td>
          </tr>
        </table></td>
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
        <td class="subtitulob">&nbsp;&nbsp;&nbsp;&nbsp;Funcionalidades a serem desativadas </td>
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
        <td width="25%" class="campob">Funcionalidade a ser desativada:</td>
        <td width="47%"><input name="txtFuncDesat" type="text" id="txtFuncDesat" value="<%=str_FuncDesat%>" size="70" maxlength="100"></td>
        <td width="12%">&nbsp;</td>
        <td width="16%">&nbsp;</td>
      </tr>
      <tr>
        <td class="campob">Data desativa&ccedil;&atilde;o:</td>
        <td><table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="31%"><input name="txtDtDesliga" type="text" class="txtCampo" id="txtDtDesliga" value="<%=dat_DtDesliga%>" size="12" maxlength="12" readonly>
              <a href="javascript:show_calendar(true,'frm_Cad_PDS_Func.txtDtDesliga','DD/MM/YYYY')"><img src="../../../imagens/show-calendar.gif" width="24" height="22" border="0"></a></td>
            <td width="32%" class="campob"><div align="right">Hora desativa&ccedil;&atilde;o: </div></td>
            <td width="37%"><input name="txtHrDesliga" type="text" class="txtCampo" id="txtHrDesliga" value="<%=hor_HrDesliga%>" size="2" maxlength="2" onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name,this);" >
              <span class="style3">              :</span>              <input name="txtMnDesliga" type="text" class="txtCampo" id="txtMnDesliga" value="<%=hor_MnDesliga%>" size="2" maxlength="2" onKeyUp="javascript: Verifica_Dif_Numeros(this.value,this.name);" > 
              <span class="campo">HH:MM</span></td>
          </tr>
        </table></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td valign="top" class="campob">Procedimento de desligamento:</td>
        <td><textarea name="txtProcDesl" cols="50" rows="5" id="txtProcDesl" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=str_ProcDesl%></textarea></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  
	  <tr> 
		<td>&nbsp;</td>		
		<td>
			<font face="Verdana" size="2"><font size="1">Caracteres digitados&nbsp;</font><b> 
		  	<input type="text" name="txttamanhoProcDesl" size="5" value="0" maxlength="50" readonly>
		  	</b></font><font face="Verdana" size="1">(Máximo 300 caracteres)</font> 
		</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	
      <tr>
        <td valign="top"> <table width="90%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td class="campob">Destinos dos dados pr&oacute;prios:</td>
            </tr>
            <tr>
              <td class="campob"><div align="center">Tempo de reten&ccedil;&atilde;o </div></td>
            </tr>
          </table> </td>
        <td><textarea name="txtDestDados" cols="50" rows="5" id="txtDestDados" onKeyUp="javascript: VerifiCacaretersEspeciais(this.value,this.name);pega_tamanho(this.name);"><%=str_DestDados%></textarea></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  
	  <tr> 
		<td>&nbsp;</td>		
		<td>
			<font face="Verdana" size="2"><font size="1">Caracteres digitados&nbsp;</font><b> 
		  	<input type="text" name="txttamanhoDestDados" size="5" value="0" maxlength="50" readonly>
		  	</b></font><font face="Verdana" size="1">(Máximo 300 caracteres)</font> 
		</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	
      <tr>
        <td>&nbsp;</td>
        <td><input type="hidden" value="<%=int_IdTaskProject%>" name="idTaskProject">
          <input type="hidden" value="<%=int_Plano%>" name="pintPlano">
        <input type="hidden" value="<%=int_CdSeqFunc%>" name="pCdSeqFunc">
        <input type="hidden" value="<%=str_Acao%>" name="pAcao"></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table>
    <table width="625" border="0" align="center">
      <tr>
        <td width="85"><a href="javascript:confirma_pds_sub()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
        <td width="23"><b></b></td>
        <td width="187">
          <%if str_Acao = "A" then%>
          <a href="javascript:confirma_Exclusao();"><img src="../img/botao_excluir.gif" width="85" height="19" border="0"></a>
        <%end if%></td>
        <td width="173"><%if str_Acao = "C" then%>
          <a href="javascript:history.go(-1)"><img src="../img/botao_voltar.gif" width="85" height="19" border="0"></a>
          <%end if%></td>
        <td width="10"></td>
        <td width="9"></td>
        <td width="8">&nbsp;</td>
        <td width="106"><div align="center"></div></td>
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
