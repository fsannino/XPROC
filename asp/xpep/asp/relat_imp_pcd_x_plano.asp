<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Expires=0

IF request("str_Tipo_Saida")="Excel" THEN
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
END IF

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

'response.Write "<br><br><br>int_Onda " &  int_Onda & "<br>"
'response.Write "int_Fase " &  int_Fase & "<br>"
'response.Write "int_Plano " &  int_Plano & "<br>"
'response.Write "int_Atividade " &  int_Atividade & "<br>"
'response.End()

if strPlano <> "" then
	vetPlano = split(strPlano,"|")
	int_Plano = vetPlano(0)
else
	int_Plano = 0
end if

if strAtividade <> "" then
	vetAtividade = split(strAtividade,"|")
	int_Atividade = vetAtividade(0)
else
	int_Atividade = 0
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
sqlPlano = sqlPlano & " FROM ONDA, XPEP_PLANO_ENT_PRODUCAO ENT_PROD, XPEP_PLANO_TAREFA_PCD PCD"
sqlPlano = sqlPlano & " WHERE ONDA.ONDA_CD_ONDA = ENT_PROD.PLAN_NR_CD_ONDA"
sqlPlano = sqlPlano & " AND ENT_PROD.PLAN_NR_SEQUENCIA_PLANO = PCD.PLAN_NR_SEQUENCIA_PLANO"

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
sqlPlano = sqlPlano & " , ENT_PROD.PLAN_NR_SEQUENCIA_PLANO"

'Response.write sqlPlano
'Response.end
		
set rds_Plano = db_Cogest.Execute(sqlPlano)

sqlPlanoPCD = ""
sqlPlanoPCD = sqlPlanoPCD & "SELECT PLAN_NR_SEQUENCIA_PLANO"			
sqlPlanoPCD = sqlPlanoPCD & ", PLTA_NR_SEQUENCIA_TAREFA "		
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_SISTEMA_LEGADO "			
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_DADO_A_SER_MIGRADO "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_TIPO_ATIVIDADE "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_TIPO_DADO "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_CARAC_DADO "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_EXTRACAO "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_EXTRACAO "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_QTD_TEMPO_EXEC_CARGA "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_UNID_TEMPO_EXEC_CARGA "			
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_ARQ_CARGA "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_NR_VOLUME "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_DEPENDENCIAS "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_DT_EXTRACAO "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_DT_CARGA_INICIO "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_DT_CARGA_FIM "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_TX_COMO_EXECUTA "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_NR_ID_PLANO_CONTINGENCIA "
sqlPlanoPCD = sqlPlanoPCD & ", PPCD_NR_ID_PLANO_COMUNICACAO "				
sqlPlanoPCD = sqlPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_TEC "
sqlPlanoPCD = sqlPlanoPCD & ", USUA_CD_USUARIO_RESP_LEG_FUN "			
sqlPlanoPCD = sqlPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_TEC "	
sqlPlanoPCD = sqlPlanoPCD & ", USUA_CD_USUARIO_RESP_SIN_FUN "		
sqlPlanoPCD = sqlPlanoPCD & ", ATUA_TX_OPERACAO "
sqlPlanoPCD = sqlPlanoPCD & ", ATUA_CD_NR_USUARIO "
sqlPlanoPCD = sqlPlanoPCD & ", ATUA_DT_ATUALIZACAO "
sqlPlanoPCD = sqlPlanoPCD & " FROM XPEP_PLANO_TAREFA_PCD"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

</script>
<head>
<title></title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

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
<link href="../../../css/biblioteca.css" rel="stylesheet" type="text/css">
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript">	

</script>
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<table width="670" height="19" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="18">&nbsp;</td>
        <td width="652"><div align="right">
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></font></div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormataData(Date())%></font></div></td>
      </tr>
</table>
<table width="670"  border="0" cellspacing="0" cellpadding="1">
  <tr>
	<td width="670"><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Rela&ccedil;&atilde;o de Plano de Convers&otilde;es de Dados - PCD</font></strong></div></td>
  </tr>
  <tr>
	<td width="670"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ordenado por onda - fase - plano e data limite para comunica&ccedil;&atilde;o</font></div></td>
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
		if boo_MostraOnda or boo_MostraFase then
		%>
			<table width="670"  border="0" cellspacing="5" cellpadding="1">
			  <tr>
				<td width="136">				  
				  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
				  <strong>
				  <% if boo_MostraOnda then %>
				  Onda -<%=rds_Plano("ONDA_TX_DESC_ONDA")%>
			      <% end if %>
			    </strong>			    </font></td>
				<td width="159">
				  <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
				  <% if boo_MostraFase then %>
				Fase -<%=rds_Plano("PLAN_NR_CD_FASE")%>
				<% end if %>
			      </font></strong></td>
				<td width="349"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Plano - <%=rds_Plano("PLAN_TX_SIGLA_PLANO")%></font></strong></td>
			  </tr>
			</table>
			<%	
		end if		
		 		
		sqlPlanoPCD_Complemento = ""
		sqlPlanoPCD_Complemento = sqlPlanoPCD_Complemento & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & rds_Plano("PLAN_NR_SEQUENCIA_PLANO")		
		
		if int_Atividade <> 0 then
			sqlPlanoPCD_Complemento = sqlPlanoPCD_Complemento & " AND PLTA_NR_SEQUENCIA_TAREFA = " & int_Atividade
		end if
		
		'Response.write sqlPlanoPCD & sqlPlanoPCD_Complemento
		'Response.end
		
		set rds_PCD = db_Cogest.Execute(sqlPlanoPCD + sqlPlanoPCD_Complemento)
		%>		
		<table width="680" border="0" cellpadding="1" cellspacing="3">
		<% if boo_Mostra_Cabec = True then %>	
		<% if request("str_Tipo_Saida") <> "Excel" then %>
		  <tr width="680">
		    <td colspan="6"><img src="../img/tit_tab_imp_Fundo_PCD1.gif" width="680" height="25"></td>
	      </tr>
		  <% else %>  
		  <tr bgcolor="#639ACE" width="680">			
			<td width="120" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atividade</strong></font></td>
			<td width="96" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Dado a ser Migrado</strong></font></td>
			<td width="116" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Sistema Legados de Origem</strong></font></td>
			<td width="109" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Arquivos de Carga</strong></font></td>
			<td width="102" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Depend&ecirc;ncia</strong></font></td>
			<td width="112" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Como Executa</strong></font></td>
		  </tr>		
		<% end if %>  
		<% end if %>		  
		<%
		int_TotRegistroPAC = rds_PCD.recordcount
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
			sql_TarefaProject = sql_TarefaProject & " AND TASK_UID = " & rds_PCD("PLTA_NR_SEQUENCIA_TAREFA") 			
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
			 <tr bgcolor="<%=str_Cor%>" width="680">				
				<td colspan="6" bgcolor="<%=str_Cor%>"><img src="../img/001103.gif" width="680" height="1"></td>
			  </tr>	
			<% end if %>
			  <tr bgcolor="<%=str_Cor%>" width="680">				
				<td width="120" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NomeAtividade%></font></td>
				<td width="96" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCD("PPCD_TX_DADO_A_SER_MIGRADO")%></font></td>
				<td width="116" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCD("PPCD_TX_SISTEMA_LEGADO")%></font></td>
				<td width="109" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCD("PPCD_TX_ARQ_CARGA")%></font></td>
				<td width="127" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCD("PPCD_TX_DEPENDENCIAS")%></font></td>
				<td width="112" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCD("PPCD_TX_COMO_EXECUTA")%></font></td>
			  </tr>			
			<% 
			rds_PCD.movenext
		Loop
		rds_PCD.close
		rds_Plano.movenext		
	Loop 
	rds_Plano.Close
	set rds_Plano = Nothing
	set rds_PCD = Nothing
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
