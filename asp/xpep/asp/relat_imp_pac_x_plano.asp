<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Expires=0

if request("str_Tipo_Saida")="Excel" then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

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

'response.Write(int_Onda)
'response.Write(int_Fase)
'response.Write(int_Plano)
'response.Write(int_Atividade)
'response.end 

if strPlano <> "" then
	vetPlano = split(strPlano,"|")
	int_Plano = vetPlano(0)
else
	int_Plano = 0
end if

'if strAtividade <> "" then
'	vetAtividade = split(strAtividade,"|")
'	int_Atividade = vetAtividade(0)
'else
'	int_Atividade = 0
'end if

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
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link href="../../../css/biblioteca.css" rel="stylesheet" type="text/css">
	<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
	<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
	<title></title>
	
	<script language="JavaScript" type="text/JavaScript">		
		function MM_reloadPage(init) {  //reloads the window if Nav4 resized
		  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
			document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
		  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
		}
		MM_reloadPage(true);
	</script>
</head>

<%=ls_Script%>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<table width="670" height="19" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20">&nbsp;</td>
        <td width="650"><div align="right">
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></font></div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormataData(Date())%></font></div></td>
      </tr>
</table>  
<table width="670"  border="0" cellspacing="0" cellpadding="1">     
  <tr>       
	<td width="76%" class="subtitulob"><div align="center" class="campob">Rela&ccedil;&atilde;o de Plano de Ações Corretivas / Contingências - PAC</div></td>
  </tr> 
  <tr>
	<td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ordenado por onda - fase - plano</font></div></td>
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
				<table width="670"  border="0" cellspacing="5" cellpadding="1">
				  <tr>
					<td width="150">				  
					  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
					  <strong>
					  <% if boo_MostraOnda then %>
					  Onda -<%=rds_Plano("ONDA_TX_DESC_ONDA")%>
					  <%else%>
					  &nbsp;
					  <% end if %>
					</strong>			    </font></td>
					<td width="139">
					  <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
					  <% if boo_MostraFase then %>
					  Fase -<%=rds_Plano("PLAN_NR_CD_FASE")%>
					  <% end if %>
					</font></strong></td>
					<td width="216"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Plano - <%=rds_Plano("PLAN_TX_SIGLA_PLANO")%></font></strong></td>
					<td width="122"><div align="right"></div></td>
				  </tr>
				</table>
				<%	
			end if
		 		
			sqlPlanoPAC_Complemento = ""
			sqlPlanoPAC_Complemento = sqlPlanoPAC_Complemento & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & rds_Plano("PLAN_NR_SEQUENCIA_PLANO")		
			
			if int_Atividade <> 0 then
				sqlPlanoPAC_Complemento = sqlPlanoPAC_Complemento & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_Atividade
			end if
		
			set rds_PAC = db_Cogest.Execute(sqlPlanoPAC + sqlPlanoPAC_Complemento)
			%>		
			<table width="685" border="0" cellpadding="1" cellspacing="3">
			<% if boo_Mostra_Cabec = True then %>
			<% if request("str_Tipo_Saida") <> "Excel" then %>
			  <tr>
			    <td colspan="4"><img src="../img/tit_tab_imp_Fundo_PAC1.gif" width="680" height="25"></td>
		      </tr>
			  <% else %>  
			  <tr bgcolor="#639ACE">				
				<td width="280"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atividade</strong></font></td>
				<td width="184"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Problemas</strong></font></td>
				<td width="130"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ações Corretivas/Contingênciais</strong></font></td>
				<td width="71"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data Aprovação</strong></font></td>
			  </tr>			
			  <% end if %> 
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
				<% if request("str_Tipo_Saida") <> "Excel" then %>
			  <tr width="700" bgcolor="<%=str_Cor%>">			  		  	
				<td colspan="4" bgcolor="<%=str_Cor%>"><img src="../img/001103.gif" width="100%" height="1"></td>
			  </tr>
  				<% end if %>
			  <tr bgcolor="<%=str_Cor%>">				
				<td width="280" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NomeAtividade%></font></td>
				<td width="184" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PAC("PPAC_TX_PROBLEMAS")%></font></td>
				<td width="130" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PAC("PPAC_TX_ACOES_CORR_CONT")%></font></td>
				<td width="71" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormataData(rds_PAC("PPAC_DT_APROVACAO"))%></font></td>
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
<table width="680"  border="0" cellspacing="0" cellpadding="1">
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

<%
db_Cronograma.close
set db_Cronograma = nothing

db_Cogest.close
set db_Cogest = nothing
%>
</body>
<script language="javascript">	
	function fechar()
	{
		window.top.close();	
	}	
		
	setTimeout('fechar()',1);	
	window.top.frame2.focus();
	window.top.frame2.print();
</script>
</html>
