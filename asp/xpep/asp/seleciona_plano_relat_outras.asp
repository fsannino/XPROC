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

if Request("pTpRel") <> "" then
	str_TpRel=Request("pTpRel")
else
	str_TpRel=Request("hidTpRel")
end if

if Trim(request("pSiglaPlano")) = "PAI" then
	str_Titulo = "Plano de Acionamento de Interface e Processo Batch - PAI"
	str_Sigla_Plano = Trim(request("pSiglaPlano"))
elseif Trim(request("pSiglaPlano")) = "PDS" then
	str_Titulo = "Plano de Desligamento de Sistemas Legados - PDS"
	str_Sigla_Plano = Trim(request("pSiglaPlano"))
elseif Trim(request("pSiglaPlano")) = "PCD" then
	str_Titulo = "Plano de Conversão de Dados - PCD"
	str_Sigla_Plano = Trim(request("pSiglaPlano"))
elseif Trim(request("pSiglaPlano")) = "PPO" then
	str_Titulo = "Plano de Parada Operacional - PPO"
	str_Sigla_Plano = Trim(request("pSiglaPlano"))
end if

if Request("selOnda") <> "" then	
	str_Cd_Onda = Request("selOnda")
else
	str_CD_Onda = 0
end if

if Request("selFases") <> "" then	
	str_Cd_Fases = Request("selFases")
else
	str_Cd_Fases = 0
end if

if Request("selTask1") <> "" then	
	str_Cd_Task1 = Request("selTask1")
else
	str_Cd_Task1 = 0
end if

'response.Write("Onda: " & str_CD_Onda)
'response.Write("Fase: " & str_Cd_Fases)
'response.Write("Task: " & str_Cd_Task1)
'response.Write("Plano: " & str_Sigla_Plano)
'response.End()

' ================= PREPARA COMBO DE ONDA
str_Sql_Onda = ""
str_Sql_Onda = str_Sql_Onda & " SELECT ONDA_TX_DESC_ONDA "
str_Sql_Onda = str_Sql_Onda & " , ONDA_CD_ONDA, ONDA_TX_ABREV_ONDA "
str_Sql_Onda = str_Sql_Onda & " FROM ONDA "
str_Sql_Onda = str_Sql_Onda & " WHERE "
str_Sql_Onda = str_Sql_Onda & " ONDA_CD_ONDA<>4 "
str_Sql_Onda = str_Sql_Onda & " ORDER BY ONDA_TX_DESC_ONDA"
'response.Write(str_Sql_Onda)
'response.End()
set rds_onda = db_Cogest.execute(str_Sql_Onda)


' ================= PREPARA COMBO DE TAREFAS
str_Sql_Plano = ""
str_Sql_Plano = str_Sql_Plano & " SELECT PLAN_TX_SIGLA_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_NR_SEQUENCIA_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_TX_DESCRICAO_PLANO"
str_Sql_Plano = str_Sql_Plano & " , PLAN_TX_IDENTACAO"
str_Sql_Plano = str_Sql_Plano & " FROM XPEP_PLANO_ENT_PRODUCAO"
str_Sql_Plano = str_Sql_Plano & " WHERE PLAN_NR_SEQUENCIA_PLANO > 0 "
if str_Sigla_Plano <> "1" then
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_TX_SIGLA_PLANO = '" & str_Sigla_Plano & "' "
end if
if str_CD_Onda <> "" then
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_ONDA = " & str_CD_Onda
else
	str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_ONDA = 0"
end if

if str_CD_Onda = 5 or str_CD_Onda = 7 then
	if str_Cd_Fases <> "" then
		str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_FASE = " & str_Cd_Fases
	else	
		str_Sql_Plano = str_Sql_Plano & " AND PLAN_NR_CD_FASE = 0"
	end if
end if
'Response.WRITE ("<br>" & "<br>" & str_Sql_Plano)
'RESPONSE.END
set rds_Plano=db_Cogest.execute(str_Sql_Plano)
if not rds_Plano.EOF then
	str_Cd_Plano = rds_Plano("PLAN_NR_SEQUENCIA_PLANO")
	str_Identacao = rds_Plano("PLAN_TX_IDENTACAO")
	str_Cd_Plano2 = rds_Plano("PLAN_NR_SEQUENCIA_PLANO") & "|" & rds_Plano("PLAN_TX_IDENTACAO") & "|" & str_Sigla_Plano
else
	str_Cd_Plano = ""
	str_Identacao = ""
	str_Cd_Plano2 = ""
end if
rds_Plano.close
set rds_Plano = nothing

'response.Write("Fase: " & str_Cd_Plano)
'response.Write("Task: " & str_Identacao)

'*** PEGA NR DO PLANO DO PROJECT
str_TpPlano = ""
str_TpPlano = str_TpPlano & "Select PLAN_TX_SIGLA_PLANO, PLAN_NR_CD_PROJETO_PROJECT "
str_TpPlano = str_TpPlano & " From XPEP_PLANO_ENT_PRODUCAO "
str_TpPlano = str_TpPlano & " WHERE "
str_TpPlano = str_TpPlano & " PLAN_NR_SEQUENCIA_PLANO = " & Trim(str_Cd_Plano)
'RESPONSE.Write(str_TpPlano)
if str_Cd_Plano <> "" then
	set rdsTpPlano = db_Cogest.Execute(str_TpPlano)
	if not rdsTpPlano.Eof then
	   int_Cd_Projeto_Project2 = rdsTpPlano("PLAN_NR_CD_PROJETO_PROJECT")   
	else
	   int_Cd_Projeto_Project2 = ""
	end if
	rdsTpPlano.close
	set rdsTpPlano = Nothing
	'response.Write "int_Cd_Projeto_Project - " & int_Cd_Projeto_Project & "<br>"
	'RESPONSE.End()
end if

' ================= PREPARA COMBO DE TAREFAS
str_Sql_Task = ""
str_Sql_Task = str_Sql_Task & " SELECT   "
str_Sql_Task = str_Sql_Task & " TASK_UID"
str_Sql_Task = str_Sql_Task & " , TASK_NAME"
str_Sql_Task = str_Sql_Task & " , RESERVED_DATA"
str_Sql_Task = str_Sql_Task & " , TASK_START_DATE"
str_Sql_Task = str_Sql_Task & " , TASK_FINISH_DATE"
str_Sql_Task = str_Sql_Task & " FROM MSP_TASKS"
str_Sql_Task = str_Sql_Task & " WHERE (LEN(TASK_OUTLINE_NUM) = 11 or LEN(TASK_OUTLINE_NUM) = 12)"

if str_Identacao <> "" then
	str_Sql_Task = str_Sql_Task & " AND TASK_OUTLINE_NUM LIKE '" & TRIM(str_Identacao) & "%'"
	str_Sql_Task = str_Sql_Task & " AND PROJ_ID = " & int_Cd_Projeto_Project2
else
	str_Sql_Task = str_Sql_Task & " AND TASK_OUTLINE_NUM = '99999'"
	str_Sql_Task = str_Sql_Task & " AND PROJ_ID = 0"
end if

str_Sql_Task = str_Sql_Task & " ORDER BY TASK_NAME"

'response.Write(str_Sql_Task)
'response.End()

set rds_Task=db_Cronograma.execute(str_Sql_Task)

%>  

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>:: Cutover - Consulta das Atividades</title>
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

var int_Controle = 0
var str_Param = ""

	function chamapagina()
	{		
		//alert(int_Controle)
		if(int_Controle == 1)
		   	{
		   	str_Param = "hidTpRel="+document.frm1.hidTpRel.value+"&pSiglaPlano="+document.frm1.pSiglaPlano.value+"&selOnda="+document.frm1.selOnda.value
			//alert(str_Param)
			//+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+'&txtSubModulo='+document.frm1.txtSubModulo.value
			window.location.href='seleciona_plano_relat_outras.asp?'+str_Param
			}
		if(int_Controle == 2)
		   	{
		   	str_Param = "hidTpRel="+document.frm1.hidTpRel.value+"&pSiglaPlano="+document.frm1.pSiglaPlano.value+"&selOnda="+document.frm1.selOnda.value+'&selFases='+document.frm1.selFases.value
			//alert(str_Param)
			//+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+'&txtSubModulo='+document.frm1.txtSubModulo.value
			window.location.href='seleciona_plano_relat_outras.asp?'+str_Param
			}
			
		if(int_Controle == 3)
		   	{
		   	str_Param = "hidTpRel="+document.frm1.hidTpRel.value+"&pSiglaPlano="+document.frm1.pSiglaPlano.value+"&selOnda="+document.frm1.selOnda.value+'&selFases='+document.frm1.selFases.value+'&selTask1='+document.frm1.selTask1.value+'&selPlano='+document.frm1.selPlano.value
			//alert(str_Param)
			//+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+'&txtSubModulo='+document.frm1.txtSubModulo.value
			window.location.href='seleciona_plano_relat_outras.asp?'+str_Param
			}

	}
		
	function Confirma()
	{	
		if (document.frm1.hidTpRel.value == '1')   
	   		{	
			document.frm1.action="relat_pai_x_atividade.asp";          
			document.frm1.submit();
			}
			
		if (document.frm1.hidTpRel.value == '2')   
			{	
			document.frm1.action="relat_pds_x_atividade.asp";          
			document.frm1.submit();
			}	
		if (document.frm1.hidTpRel.value == '3')   
			{	
			document.frm1.action="relat_pcd_x_atividade.asp";          
			document.frm1.submit();
			}	
		if (document.frm1.hidTpRel.value == '4')   
			{	
			document.frm1.action="relat_ppo_x_atividade.asp";          
			document.frm1.submit();
			}	

	}
	
	function Limpa()
	{
		str_Param = "hidTpRel="+document.frm1.hidTpRel.value+"&pSiglaPlano="+document.frm1.pSiglaPlano.value;
		window.location.href='seleciona_plano_relat_outras.asp?'+str_Param
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
<form name="frm1" method="post" action="">
	<table width="98%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="10%">&nbsp;</td>
        <td width="77%">&nbsp;</td>
        <td width="13%">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td class="subtitulo"><strong>Consulta das Atividades - <%=str_Titulo%></strong></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<table width="102%"  border="0" cellspacing="10" cellpadding="1">
      <tr>
        <td width="15%">&nbsp;</td>
        <td width="77%">&nbsp;</td>
        <td width="4%">&nbsp;</td>
        <td width="4%">&nbsp;</td>
      </tr>
      <tr>
        <td><div align="right"><span class="campob">Onda:</span></div></td>
        <td><select name="selOnda" size="1" class="cmdOnda" onChange="javascript:int_Controle=1;chamapagina();">
          <option value="">== Todas as Ondas ==</option>
          <%
	do until rds_onda.EOF = True		
		if trim(rds_onda("ONDA_CD_ONDA")) <> "1" and trim(rds_onda("ONDA_CD_ONDA")) <> "2" then 		
			if TRIM(str_Cd_Onda)=trim(rds_onda("ONDA_CD_ONDA")) then%>
          <option selected value=<%=rds_onda("ONDA_CD_ONDA")%>><%=rds_onda("ONDA_TX_DESC_ONDA")%></option>
          <%else%>
          <option value=<%=rds_onda("ONDA_CD_ONDA")%>><%=rds_onda("ONDA_TX_DESC_ONDA")%></option>
          <%
			end if
		end if
	rds_onda.MOVENEXT
	loop
	rds_onda.close
	set rds_onda = Nothing
	%>
        </select>
        <input name="pSiglaPlano" type="hidden" value="<%=str_Sigla_Plano%>"></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <% if str_CD_Onda = 5 or str_CD_Onda = 7 then %>
      <tr>
        <td class="campob"><div align="right">Fase:</div></td>
        <td><select name="selFases" size="1" class="cmdOnda" onChange="javascript:int_Controle=2;chamapagina();">
          <% if str_Cd_Fases = "0" then %>
			  <option value="" selected>== Todas as Fase ==</option>
			  <option value="1">Fase 1</option>
			  <option value="2">Fase 2</option>
          <% elseif str_Cd_Fases = "1" then %>
			  <option value="0">== Todas as Fase ==</option>
			  <option value="1" selected>Fase 1</option>
			  <option value="2">Fase 2</option>
          <% elseif str_Cd_Fases = "2" then %>
			  <option value="0">== Todas as Fase ==</option>
			  <option value="1">Fase 1</option>
			  <option value="2" selected>Fase 2</option>
          <% end if %>
        </select>        </td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <% end if %>
      <tr>
        <td class="campob"><div align="right" class="campob">Plano:</div></td>
        <td><div align="right" class="campob">
          <div align="left"><span class="subtitulo"><strong><%=str_Titulo%>
            <input name="selPlano" type="hidden" id="selPlano" value="<%=str_Cd_Plano2%>">
          </strong></span></div>
        </div></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td valign="top"><div align="right" class="campob">Atividades:</div></td>
        <td>		
		<select name="selTask1" size="1" class="cmdTask"  onChange="javascript:int_Controle=3;chamapagina();">
          <option value="">== Todas as Atividades ==</option>
          <%
		Do until rds_Task.eof=true
			 If Trim(Request("selTask1")) = Trim(rds_Task("TASK_UID") & "|" & rds_Task("RESERVED_DATA")) then
		  %>
          <option selected value=<%=rds_Task("TASK_UID") & "|" & rds_Task("RESERVED_DATA")%>><%=rds_Task("TASK_NAME")%></option>
          <%else%>
          <option value=<%=rds_Task("TASK_UID") & "|" & rds_Task("RESERVED_DATA")%>><%=rds_Task("TASK_NAME")%></option>
          <%
			end if
			rds_Task.movenext
		  Loop
		  rds_Task.close
		  set rds_Task = Nothing
            	%>
        </select></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  
	  <%if Right(Trim(str_Titulo),3) = "PPO" and Request("selOnda") = "7" then%>
		<tr>
		  <td height="21" valign="top"><div align="right" class="campob">Sub-Atividade:</div></td>
		  <td><!--#include file="../includes/inc_combo_tarefas_nivel2.asp" --></td>
		  <td>&nbsp;</td>
		</tr>
		<%end if%>	
	
    </table>
    <table width="625" border="0" align="center">
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td><%'=response.write(" - Plano2: " & session("CD_Plano2"))%></td>
        <td></td>
        <td></td>
        <td>&nbsp;</td>
        <td><input name="hidTpRel" type="hidden" id="hidTpRel" value="<%=str_TpRel%>"></td>
      </tr>
      <tr>
        <td width="26"><a href="javascript:Confirma()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
        <td width="26"><b></b></td>
        <td width="195"><a href="javascript:Limpa()"><img src="../img/limpar_01.gif" width="85" height="19" border="0"></a>
            <!--<a href="javascript:Limpa();"><img src="../img/limpar_01.gif" width="85" height="19" border="0"></a>--></td>
        <td width="27"></td>
        <td width="50">&nbsp;</td>
        <td width="28"></td>
        <td width="26">&nbsp;</td>
        <td width="159"></td>
      </tr>
    </table>
	<%
	db_Cronograma.close
	set db_Cronograma = nothing
	
	db_Cogest.close
	set db_Cogest = nothing
	%>
</form>
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
