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

int_CD_Onda = request("pOnda")
int_Plano = request("pPlano")

if session("CD_Plano2") = "" then	
	int_Plano2 = request("pPlano2")
	Session("CD_Plano2") = int_Plano2
else
	vet_int_Plano2 = Split(Session("CD_Plano2"),"|")
	int_Plano2 = vet_int_Plano2(0)
end if

'Response.write "<br><br>"
'Response.write "int_CD_Onda - " & int_CD_Onda & "<br>"
'Response.write "int_Plano - " & int_Plano & "<br>"
'Response.write "int_Plano2 - " & int_Plano2 & "<br>"
'Response.end

'*********** SELECIONA SIGLA E DESCRIÇÃO DO PLANO SELECIONADO OARA A CRIAÇÃO DO PCM ****
str_SelPlano = ""
str_SelPlano = str_SelPlano & "SELECT PLAN_TX_SIGLA_PLANO, PLAN_TX_DESCRICAO_PLANO " 
str_SelPlano = str_SelPlano & " FROM XPEP_PLANO_ENT_PRODUCAO "
str_SelPlano = str_SelPlano & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & int_Plano2

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

str_PCM_Sub = ""
str_PCM_Sub = str_PCM_Sub & " SELECT  "
str_PCM_Sub = str_PCM_Sub & " PLAN_NR_SEQUENCIA_PLANO"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_ATIVIDADE"
str_PCM_Sub = str_PCM_Sub & " , PPCM_NR_SEQUENCIA_TAREFA"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_TP_COMUNICACAO"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_O_QUE_COMUNICAR"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_PARA_QUEM"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_UNID_ORGAO"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_QUANDO_OCORRE"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_RESP_CONTEUDO"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_RESP_DIVULGACAO"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_COMO"
str_PCM_Sub = str_PCM_Sub & " , PPCM_TX_APROVADOR_PB"
str_PCM_Sub = str_PCM_Sub & " , PPCM_DT_APROVACAO"
str_PCM_Sub = str_PCM_Sub & " FROM XPEP_PLANO_TAREFA_PCM "
str_PCM_Sub = str_PCM_Sub & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & int_Plano2  
'response.Write(str_PCM_Sub)
'response.End()

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
	function confirma_Exclusao(intPlano2,int_CD_Onda)
	{
		  if(confirm("Confirma a exclusão deste Registro?"))
		  {
		    document.frm_Plano_PCM.pAcao.value = 'E';
			document.frm_Plano_PCM.pintPlano2.value = intPlano2;
			document.frm_Plano_PCM.pOnda.value = int_CD_Onda;			
			document.frm_Plano_PCM.action='grava_plano.asp' 			        
			document.frm_Plano_PCM.submit();	
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
<form name="frm_Plano_PCM">
    <table width="75%" border="0" cellpadding="0" cellspacing="7">
      <tr>
        <td width="11%"><div align="right" class="subtitulob">Onda:</div></td>
        <td colspan="2" class="subtitulo"><%=str_Desc_Onda%></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td width="19%" class="subtitulob">Plano Origem:</td>
        <td width="70%" class="subtitulo"><%=str_Plano%></td>
      </tr>
    </table>   
    <table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="2" bgcolor="#CCCCCC"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center" class="subtitulob">Plano de Comunica&ccedil;&atilde;o - PCM </div></td>
      </tr>
      <tr>
        <td>	
			<input type="hidden" value="<%=str_Acao%>" name="pAcao">
            <input type="hidden" value="<%=int_Plano2%>" name="pintPlano2">
            <input type="hidden" value="<%=int_Id_TarefaProject%>" name="idTaskProject2">
            <input type="hidden" value="PCM" name="pPlano">
			<input type="hidden" value="" name="pOnda">
            <input type="hidden" value="0" name="pCdSeqPCM"></td>
      </tr>
    </table>
    <br>
    <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
      <tr bgcolor="#CCCCCC"> 
        <td width="7%" bgcolor="#9C9A9C" class="titcoltabela"><a href="inclui_altera_plano_pcm_sub.asp?pAcao=I&pPlano=<%=int_Plano%>&pPlano2=<%=int_Plano2%>&pOnda=<%=int_CD_Onda%>&pPlano_Origem=<%=str_Plano%>"><img src="../img/botao_novo_off_02.gif" alt="Incluir uma nova Atividade para o Plano de Comunicação" width="34" height="23" border="0"></a></td>
        <td colspan="6"><div align="center"></div>          
        <div align="center" class="campob">Equipe respons&aacute;vel pela implementa&ccedil;&atilde;o</div>          <div align="center"></div></td>
      </tr>
      <tr bgcolor="#CCCCCC">
        <td bgcolor="#9C9A9C" class="titcoltabela">&nbsp;</td>
        <td width="10%" class="titcoltabela"><div align="center"><span class="campob">Atividade</span></div></td>
        <td width="10%" class="titcoltabela"><div align="center">Tipo de Comunica&ccedil;&atilde;o</div></td>
        <td width="27%" class="titcoltabela"><div align="center">O que comunicar</div></td>
        <td width="20%" class="titcoltabela"><div align="center">Para quem </div></td>
        <td width="15%" class="titcoltabela"><div align="center">Unidade/&Oacute;rg&atilde;o</div></td>
        <td width="11%" class="titcoltabela"><div align="center">Dt Limite para comunica&ccedil;&atilde;o</div></td>
      </tr>
      <%
	'Response.write str_PCM_Sub 
	'Response.end
	  
	set rdsPCM_Sub = db_Cogest.Execute(str_PCM_Sub)
	if not rdsPCM_Sub.EOF then 
	      Do while not rdsPCM_Sub.EOF
	%>
      <tr bgcolor="#E9E9E9">
        <td bgcolor="#9C9A9C"><a href="inclui_altera_plano_pcm_sub.asp?pAcao=A&pPlano=<%=int_Plano%>&pPlano2=<%=int_Plano2%>&pOnda=<%=int_CD_Onda%>&pCdSeqPCM=<%=rdsPCM_Sub("PPCM_NR_SEQUENCIA_TAREFA")%>&pPlano_Origem=<%=str_Plano%>"><img src="../img/botao_abrir_off_02.gif" alt="Alterar Atividade do Plano de Comunicação" width="34" height="23" border="0"></a><a href="javascript:document.frm_Plano_PCM.pCdSeqPCM.value='<%=rdsPCM_Sub("PPCM_NR_SEQUENCIA_TAREFA")%>';confirma_Exclusao('<%=int_Plano2%>','<%=int_CD_Onda%>');"><img src="../img/botao_deletar_off_02.gif" alt="Excluir Atividade do Plano de Comunicação" width="34" height="23" border="0"></a></td>
        <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPCM_Sub("PPCM_TX_ATIVIDADE")%></div></td>
        <% 	
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtAprov = split(Trim(rdsPCM_Sub("PPCM_TX_QUANDO_OCORRE")),"/")						
		strDia = trim(vetDtAprov(1))
		if cint(strDia) < 10 then
			strDia = "0" & strDia
		end if			
		strMes = trim(vetDtAprov(0))			
		if cint(strMes) < 10 then
			strMes = "0" & strMes
		end if
		strAno = trim(vetDtAprov(2))
		dat_DtAprov = strDia & "/" & strMes & "/" & strAno 
		%>
        <td bgcolor="#FFFFFF" class="campotabela">
			<%	
			strTipoComunicacao	 = ""	
			if trim(rdsPCM_Sub("PPCM_TX_TP_COMUNICACAO")) = "INT" then
				strTipoComunicacao = "Interno"
			elseif trim(rdsPCM_Sub("PPCM_TX_TP_COMUNICACAO")) = "EXT" then
				strTipoComunicacao = "Externa"
			elseif trim(rdsPCM_Sub("PPCM_TX_TP_COMUNICACAO")) = "AMB" then
				strTipoComunicacao = "Interno/Externa"
			end if
			%>
			<div align="center"><%=strTipoComunicacao%></div>
		</td>
        <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPCM_Sub("PPCM_TX_O_QUE_COMUNICAR")%></div></td>
        <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPCM_Sub("PPCM_TX_PARA_QUEM")%></div></td>
        <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=rdsPCM_Sub("PPCM_TX_UNID_ORGAO")%></div></td>
        <td bgcolor="#FFFFFF" class="campotabela"><div align="center"><%=dat_DtAprov%></div></td>
      </tr>
      <%       rdsPCM_Sub.movenext 
	     Loop 
	end if	
	rdsPCM_Sub.close
	set rdsPCM_Sub = nothing
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
