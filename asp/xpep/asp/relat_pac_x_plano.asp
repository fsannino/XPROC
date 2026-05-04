<%
Response.Expires=0
Response.Buffer = True

on error resume next
	set db_Cogest = Server.CreateObject("ADODB.Connection")
	db_Cogest.Open Session("Conn_String_Cogest_Gravacao")
	db_Cogest.cursorlocation = 3
	
	set db_Cronograma = Server.CreateObject("ADODB.Connection")
	db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

if err.number <> 0 then		
	strMSG = "Ocorreu algum problema com o servidor!"
	Response.redirect "msg_geral.asp?pMsg=" & strMsg & "&pErroServidor=S"
end if	

int_Onda = request("selOnda")
int_Fase = request("selFases")
strPlano = request("selPlano")
strAtividade = request("selTask1")
strAtividadeSub = request("selTaskSub")

'response.Write "<br><br><br>int_Onda " &  int_Onda & "<br>"
'response.Write "int_Fase " &  int_Fase & "<br>"
'response.Write "int_Plano " &  int_Plano & "<br>"
'response.Write "int_Atividade " &  int_Atividade & "<br>"
'response.Write "strAtividadeSub " &  strAtividadeSub & "<br>"
'response.End()

if strPlano <> "" then
	vetPlano = split(strPlano,"|")
	int_Plano = vetPlano(0)
else
	int_Plano = 0
end if

if int_Onda = "7" and strAtividadeSub <> "" then
	vetAtividadeSub = split(strAtividadeSub,"|")
	int_AtividadeSub = vetAtividadeSub(0)		
			
	strIdAtividade = split(strAtividade, "|")
	int_Atividade = strIdAtividade(0)		
else
	if strAtividade <> "" then
		strValorSelTask1 = split(strAtividade, "|")
		int_Atividade = strValorSelTask1(0)		
	else
		int_Atividade = 0		
	end if
	int_AtividadeSub = 0
end if

'Response.Write "<br><br><br><br> int_AtividadeSub" & int_AtividadeSub & "<br>"
'Response.Write "int_Atividade " & int_Atividade & "<br>"
'Response.Write "intIdAtividade " &  intIdAtividade & "<br>"

Function FormataData(str_Data)
	strDia = ""		
	strMes = ""
	strAno = ""
	vet_Data = split(Trim(str_Data),"/")							
	strDia = trim(vet_Data(1))
	if cint(strDia) < 10 then
		strDia = "0" & strDia
	end if			
	strMes = trim(vet_Data(0))			
	if cint(strMes) < 10 then
		strMes = "0" & strMes
	end if
	strAno = trim(vet_Data(2))
	str_data = strDia & "/" & strMes & "/" & strAno
	FormataData = str_data	
end function

sqlPlano = ""
sqlPlano = sqlPlano & "SELECT DISTINCT ENT_PROD.PLAN_NR_SEQUENCIA_PLANO"
sqlPlano = sqlPlano & ", ENT_PROD.PLAN_TX_SIGLA_PLANO"
sqlPlano = sqlPlano & ", ENT_PROD.PLAN_TX_DESCRICAO_PLANO"
sqlPlano = sqlPlano & ", ENT_PROD.PLAN_NR_CD_FASE"
sqlPlano = sqlPlano & ", ENT_PROD.PLAN_NR_CD_ONDA"
sqlPlano = sqlPlano & ", ENT_PROD.PLAN_NR_CD_PROJETO_PROJECT"
sqlPlano = sqlPlano & ", ONDA.ONDA_TX_DESC_ONDA"
sqlPlano = sqlPlano & " FROM ONDA, XPEP_PLANO_ENT_PRODUCAO ENT_PROD, XPEP_PLANO_TAREFA_PAC PAC"
sqlPlano = sqlPlano & " WHERE ONDA.ONDA_CD_ONDA = ENT_PROD.PLAN_NR_CD_ONDA"
sqlPlano = sqlPlano & " AND ENT_PROD.PLAN_NR_SEQUENCIA_PLANO = PAC.PLAN_NR_SEQUENCIA_PLANO"

if int_Onda <> "" then
	sqlPlano = sqlPlano & " AND ENT_PROD.PLAN_NR_CD_ONDA = " & int_Onda
end if

if int_Fase <> "" then
	sqlPlano = sqlPlano & " AND ENT_PROD.PLAN_NR_CD_FASE = " & int_Fase
end if

if int_Plano <> 0 then
	sqlPlano = sqlPlano & " AND ENT_PROD.PLAN_NR_SEQUENCIA_PLANO = " & int_Plano
end if

sqlPlano = sqlPlano & " ORDER BY "
sqlPlano = sqlPlano & " ONDA.ONDA_TX_DESC_ONDA"
sqlPlano = sqlPlano & " , ENT_PROD.PLAN_NR_CD_FASE"
sqlPlano = sqlPlano & " , ENT_PROD.PLAN_TX_SIGLA_PLANO"

set rds_Plano = db_Cogest.Execute(sqlPlano)

sqlPlanoPAC = ""
sqlPlanoPAC = sqlPlanoPAC & "SELECT PLAN_NR_SEQUENCIA_PLANO"			
sqlPlanoPAC = sqlPlanoPAC & ", PLTA_NR_SEQUENCIA_TAREFA "	
sqlPlanoPAC = sqlPlanoPAC & ", PPAC_TX_PROBLEMAS "
sqlPlanoPAC = sqlPlanoPAC & ", PPAC_TX_ACOES_CORR_CONT "
sqlPlanoPAC = sqlPlanoPAC & ", PPAC_DT_APROVACAO "		
sqlPlanoPAC = sqlPlanoPAC & ", USUA_CD_USUARIO_RESP_TRAT_PROC "
sqlPlanoPAC = sqlPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_TEC "				
sqlPlanoPAC = sqlPlanoPAC & ", USUA_CD_USUARIO_RESP_SIN_FUN "
sqlPlanoPAC = sqlPlanoPAC & " FROM XPEP_PLANO_TAREFA_PAC"						
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>:: Cutover - Plano PAC</title>
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
		function impressao() 
		{		
			window.open('impressao.asp?par_PaginaPrint=relat_imp_pac_x_plano.asp&selOnda=<%=int_Onda%>&selFases=<%=int_Fase%>&selPlano=<%=int_Plano%>&selTask1=<%=int_Atividade%>&selTaskSub=<%=int_AtividadeSub%>','jan1','toolbar=no, location=no, scrollbars=no, status=no, directories=no, resizable=no, menubar=no, fullscreen=no, height=50, width=250, top=200, left=260');
		}
		
		function exporta() 
		{			
			window.open('relat_imp_pac_x_plano.asp?str_Tipo_Saida=Excel&selOnda=<%=int_Onda%>&selFases=<%=int_Fase%>&selPlano=<%=int_Plano%>&selTask1=<%=int_Atividade%>&selTaskSub=<%=int_AtividadeSub%>','jan1','toolbar=yes, location=no, scrollbars=yes, status=no, directories=no, resizable=yes, menubar=yes, fullscreen=no, height=400, width=500, status=no, top=100, left=160');
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
	<table width="800"  border="0" cellspacing="0" cellpadding="1">     
      <tr>       
        <td width="76%" class="subtitulob"><div align="center" class="campob">Rela&ccedil;&atilde;o de Plano de Ações Corretivas / Contingências - PAC</div></td>
        <td width="14%">&nbsp;</td>
      </tr> 
	  <tr>
        <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ordenado por onda - fase - plano</font></div></td>
        <td>&nbsp;</td>
      </tr>    
    </table>
	<% 		
	int_LoopPlano = 0
	int_TotRegistroPlano = rds_Plano.recordcount
	if 	int_TotRegistroPlano > 0 then
		str_Onda_Atual = ""
		str_Fase_Atual = ""
		Do until int_TotRegistroPlano = int_LoopPlano 
			int_LoopPlano = int_LoopPlano + 1				
			intProjProject = rds_Plano("PLAN_NR_CD_PROJETO_PROJECT")
			str_Cor = "#FFFFFF"
			boo_Mostra_Cabec = False
			boo_MostraOnda = False
			if str_Onda_Atual <> rds_Plano("ONDA_TX_DESC_ONDA") then
				str_Onda_Atual = rds_Plano("ONDA_TX_DESC_ONDA")
				str_Fase_Atual = ""
				boo_MostraOnda = True
				boo_Mostra_Cabec = True				
			end if
			boo_MostraFase = False
			if str_Fase_Atual <> rds_Plano("PLAN_NR_CD_FASE") then
				str_Fase_Atual = rds_Plano("PLAN_NR_CD_FASE")
				boo_MostraFase = True
			end if		
			
			boo_MostraPlano = False
			if str_Plano_Atual <> rds_Plano("PLAN_NR_SEQUENCIA_PLANO") then
				str_Plano_Atual = rds_Plano("PLAN_NR_SEQUENCIA_PLANO")
				boo_MostraPlano = True
			end if
			
			if boo_MostraPlano or boo_MostraOnda or boo_MostraFase then
				%>
				<table width="800"  border="0" cellspacing="5" cellpadding="1">
				  <tr width="800">
					<td width="201">				  
					  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
					  <strong>
					  <% if boo_MostraOnda then %>
					  Onda -<%=rds_Plano("ONDA_TX_DESC_ONDA")%>
					   <%else%>
					  &nbsp;
					  <% end if %>
					</strong>			    </font></td>
					<td width="223">
					  <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
					  <% if boo_MostraFase then %>
					  Fase -<%=rds_Plano("PLAN_NR_CD_FASE")%>
					  <% end if %>
					</font></strong></td>
					<td width="189"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Plano - <%=rds_Plano("PLAN_TX_SIGLA_PLANO")%></font></strong></td>
					<td width="154"><div align="right"></div></td>
				  </tr>
				</table>
				<%	
			end if
		 		
			sqlPlanoPAC_Complemento = ""
			sqlPlanoPAC_Complemento = sqlPlanoPAC_Complemento & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & rds_Plano("PLAN_NR_SEQUENCIA_PLANO")		
						
			if right(strPlano,3) = "PPO" then			
				if int_Atividade <> 0 then
					sqlPlanoPAC_Complemento = sqlPlanoPAC_Complemento & " AND PPAC_NR_ID_ATIVIDADE_PPO = " & int_Atividade
				end if
				
				if int_AtividadeSub <> 0 then
					sqlPlanoPAC_Complemento = sqlPlanoPAC_Complemento & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_AtividadeSub
				end if
			else
				if int_Atividade <> 0 then
					sqlPlanoPAC_Complemento = sqlPlanoPAC_Complemento & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_Atividade
				end if
			end if
					
			'Response.write "<br>" & sqlPlanoPAC + sqlPlanoPAC_Complemento & "<br>"
			'Response.end		
					
			set rds_PAC = db_Cogest.Execute(sqlPlanoPAC + sqlPlanoPAC_Complemento)
			%>		
			<table width="800" border="0" cellpadding="1" cellspacing="3">
			<% if boo_Mostra_Cabec = True then %>	
			  <tr bgcolor="#000099">
				<td width="40" bgcolor="#FFFFFF"></td>
				<td width="279" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atividade</strong></font></td>
				<td width="184" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Problemas</strong></font></td>
				<td width="194" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ações Corretivas/Contingênciais</strong></font></td>
				<td width="75" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data Aprovação</strong></font></td>
			  </tr>			
			  <% end if %>
			<%
			int_TotRegistroPAC = rds_PAC.recordcount
			int_LoopPAC = 0
			do until int_TotRegistroPAC = int_LoopPAC
				int_LoopPAC = int_LoopPAC + 1				
				
				'==================================================================================
				'==== ENCONTRA DADOS ADIOCIONAIS DA TAREFA ========================================
				sql_TarefaProject = ""
				sql_TarefaProject = sql_TarefaProject & " SELECT   "
				sql_TarefaProject = sql_TarefaProject & " TASK_UID"
				sql_TarefaProject = sql_TarefaProject & " ,TASK_NAME"
				sql_TarefaProject = sql_TarefaProject & " ,RESERVED_DATA"
				sql_TarefaProject = sql_TarefaProject & " ,TASK_START_DATE"
				sql_TarefaProject = sql_TarefaProject & " ,TASK_FINISH_DATE"
				sql_TarefaProject = sql_TarefaProject & " FROM MSP_TASKS"
				sql_TarefaProject = sql_TarefaProject & " WHERE PROJ_ID = " & intProjProject
				sql_TarefaProject = sql_TarefaProject & " AND TASK_UID = " & rds_PAC("PLTA_NR_SEQUENCIA_TAREFA") 
				
				'Response.write sql_TarefaProject & "<br>"
				'Response.end
							
				set rds_TarefaProject = db_Cronograma.Execute(sql_TarefaProject)
				
				if not rds_TarefaProject.Eof then
				   dat_Dt_Inicio = rds_TarefaProject("TASK_START_DATE")
				   dat_Dt_Termino = rds_TarefaProject("TASK_FINISH_DATE")   
				   str_NomeAtividade = rds_TarefaProject("TASK_NAME")	
				end if			
				rds_TarefaProject.close
				set rds_TarefaProject = nothing
							
				if str_Cor = "#EEEEEE" then
					str_Cor = "#FFFFFF"
				else
					str_Cor = "#EEEEEE"
				end if
				%>		
			  <tr width="800" bgcolor="<%=str_Cor%>">	
			  	<td colspan="1" width="40" bgcolor="#FFFFFF"></td>		  	
				<td colspan="4" bgcolor="<%=str_Cor%>"><img src="../img/001103.gif" width="100%" height="1"></td>
			  </tr>
  			
			  <tr bgcolor="<%=str_Cor%>">
				<td width="40" bgcolor="#FFFFFF"></td>
				<td width="279" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="inclui_altera_plano_pac.asp?pAcao=C&pCdProjProject=<%=intProjProject%>&pTArefa=<%=rds_PAC("PLTA_NR_SEQUENCIA_TAREFA")%>&pOnda=<%=rds_Plano("PLAN_NR_CD_ONDA")%>&pPlano=<%=rds_Plano("PLAN_NR_SEQUENCIA_PLANO")%>&pFase=<%=rds_Plano("PLAN_NR_CD_FASE")%>&pPlanoOriginal=<%=rds_Plano("PLAN_TX_DESCRICAO_PLANO")%>" class="link"><%=str_NomeAtividade%></a></font></td>
				<td width="184" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PAC("PPAC_TX_PROBLEMAS")%></font></td>
				<td width="194" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PAC("PPAC_TX_ACOES_CORR_CONT")%></font></td>
				<td width="75" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormataData(rds_PAC("PPAC_DT_APROVACAO"))%></font></td>
			  </tr>				
				<% 
				rds_PAC.movenext
			Loop
			rds_PAC.close
			rds_Plano.movenext		
		Loop 
		rds_Plano.Close
		set rds_Plano = Nothing
		set rds_PAC = Nothing
		str_Msg = ""	
	else
		str_Msg = "N&atilde;o existem registros para esta condi&ccedil;&atilde;o."
	end if	
	%>
</table>
	<%
	if str_Msg <> "" then 
	%>
    <table width="800"  border="0" cellspacing="0" cellpadding="1">
	  <% For i=1 to 5 %>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <% next %>
      <tr>
        <td width="146">&nbsp;</td>
        <td width="634"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Msg%></font></div></td>
        <td width="207">&nbsp;</td>
      </tr>
	  <% For j=1 to 2 %>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <% next %>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"></div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"></div></td>
        <td>&nbsp;</td>
      </tr>
	  <% For j=1 to 3 %>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
	  <% next %>	  
    </table>
	<% end if %>
	<table width="800"  border="0" cellspacing="0" cellpadding="1">
  <tr>
    <td width="155">&nbsp;</td>
    <td width="156"><div align="center"></div></td>
    <td width="122">&nbsp;</td>
    <td width="359">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="center"><a href="javascript:history.go(-1)"><img src="../img/botao_voltar.gif" alt=":: Volta tela anterior" width="85" height="19" border="0"></a></div></td>
    <td><a href="#"><img src="../img/botao_imprimir.gif" alt=":: Imprime formato relatório" width="85" height="19" border="0" onclick="impressao();"></a></td>
    <td><a href="#"><img src="../img/botao_exportar_excel.gif" alt=":: Exporta formato Excel" width="85" height="19" border="0" onclick="exporta();"></a></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>

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
