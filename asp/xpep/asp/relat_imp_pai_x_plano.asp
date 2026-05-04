<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
Response.Expires=0

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

if request("str_Tipo_Saida") = "Excel" then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

int_Onda = request("selOnda")
int_Fase = request("selFases")
int_Plano = request("selPlano")
int_Atividade = request("selTask1")

'response.Write(int_Onda)
'response.Write(int_Fase)
'response.Write(int_Plano1)
'response.Write(int_Plano2)
'response.Write(int_Atividade)

if int_Plano <> "" then	
	vet_int_Plano = Split(int_Plano,"|")
	int_Plano = vet_int_Plano(0)
end if

if int_Atividade <> "" then	
	vet_int_Atividade = Split(int_Atividade,"|")
	int_Atividade = vet_int_Atividade(0)
end if

str_SQL = ""
str_SQL = str_SQL & " SELECT  "
str_SQL = str_SQL & " dbo.XPEP_PLANO_TAREFA_PAI.PLAN_NR_SEQUENCIA_PLANO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PLTA_NR_SEQUENCIA_TAREFA "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_CD_INTERFACE"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_GRUPO "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_TIPO_PROCESSAMENTO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_NOME_INTERFACE "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_PROGRAMA_ENVOLVIDO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_PRE_REQUISITO "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_RESTRICAO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_DEPENDENCIA "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_RESP_ACIONAMENTO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_DT_INICIO "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_TX_PROCEDIMENTO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_NR_ID_PLANO_CONTINGENCIA "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PAI.PPAI_NR_ID_PLANO_COMUNICACAO"
str_SQL = str_SQL & " , dbo.ONDA.ONDA_TX_DESC_ONDA "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_TX_SIGLA_PLANO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_GERAL.PLTA_NR_ID_TAREFA_PROJECT"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_GERAL.PLTA_TX_DESC_ATIVIDADE"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_ONDA"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_FASE"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_PROJETO_PROJECT"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_TX_DESCRICAO_PLANO"
str_SQL = str_SQL & " FROM dbo.XPEP_PLANO_TAREFA_PAI INNER JOIN"
str_SQL = str_SQL & " dbo.XPEP_PLANO_TAREFA_GERAL ON "
str_SQL = str_SQL & " dbo.XPEP_PLANO_TAREFA_PAI.PLAN_NR_SEQUENCIA_PLANO = dbo.XPEP_PLANO_TAREFA_GERAL.PLAN_NR_SEQUENCIA_PLANO AND "
str_SQL = str_SQL & " dbo.XPEP_PLANO_TAREFA_PAI.PLTA_NR_SEQUENCIA_TAREFA = dbo.XPEP_PLANO_TAREFA_GERAL.PLTA_NR_SEQUENCIA_TAREFA INNER JOIN"
str_SQL = str_SQL & " dbo.XPEP_PLANO_ENT_PRODUCAO ON "
str_SQL = str_SQL & " dbo.XPEP_PLANO_TAREFA_GERAL.PLAN_NR_SEQUENCIA_PLANO = dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO INNER JOIN"
str_SQL = str_SQL & " dbo.ONDA ON dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_ONDA = dbo.ONDA.ONDA_CD_ONDA"
str_SQL = str_SQL & " WHERE dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO > 0 "
if int_Onda <> "" then
	str_SQL = str_SQL & " AND dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_ONDA = " & int_Onda
end if
if int_Fase <> "" then
	str_SQL = str_SQL & " AND dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_FASE = " & int_Fase
end if
if int_Plano1 <> "" then
	str_SQL = str_SQL & " AND dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO = " & int_Plano1
end if
if int_Atividade <> "" then
	str_SQL = str_SQL & " AND dbo.XPEP_PLANO_TAREFA_PAI.PLTA_NR_SEQUENCIA_TAREFA = " & int_Atividade
end if

str_SQL = str_SQL & " ORDER BY "
str_SQL = str_SQL & " dbo.ONDA.ONDA_TX_DESC_ONDA"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_FASE"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO"

'response.Write(str_SQL)
'response.End()

set rds_Plano = db_Cogest.Execute(str_SQL)

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

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript" type="text/JavaScript">
	function MM_reloadPage(init) {  //reloads the window if Nav4 resized
	  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
		document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
	  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
	}
	MM_reloadPage(true);
</script>
<head>	
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">	
	<link href="../../../css/biblioteca.css" rel="stylesheet" type="text/css">
	<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
	<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
	<title></title>	
</head>

<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<table width="680" height="19" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20">&nbsp;</td>
        <td width="760"><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td class="style8"><div align="right"><%=FormataData(Date())%></div></td>
      </tr>
</table>    
	<table width="680"  border="0" cellspacing="0" cellpadding="1">     
	  <tr>       
		<td width="680" class="subtitulob"><div align="center" class="campob">Relat&oacute;rio do Plano de Acionamento de Interface e Processo Batch - PAI</div></td>
	  </tr> 
	  <tr>
		<td width="680"s><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ordenado por onda - fase - plano</font></div></td>
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
			boo_MostraOnda = False
			'str_Cor = "#FFFFFF"			
			boo_Mostra_Cabec = False
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
				<table width="680"  border="0" cellspacing="5" cellpadding="1">
				  <tr>
					<td width="185">				  
					  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
					  <strong>
					  <% if boo_MostraOnda then %>
					  Onda -<%=rds_Plano("ONDA_TX_DESC_ONDA")%>
					  <% end if %>
					</strong>			    </font></td>
					<td width="134">
					  <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
					  <% if boo_MostraFase then %>
					Fase -<%=rds_Plano("PLAN_NR_CD_FASE")%>
					<% end if %>
				    </font></strong></td>
					<td width="113"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Plano - <%=rds_Plano("PLAN_TX_SIGLA_PLANO")%></font></strong></td>
					<td width="215">&nbsp;</td>
				  </tr>
</table>
		<%end if%>						
		<table width="680" border="0" cellpadding="1" cellspacing="2" bordercolor="#CCCCCC">
	<% if boo_Mostra_Cabec = True then%>	
	<% if request("str_Tipo_Saida") <> "Excel" then %>
      <tr width="680">       
        <td colspan="5"><img src="../img/tit_tab_imp_Fundo_PAI1.gif" width="680" height="25"></td>
      </tr>
    <% else %> 		
		  <tr width="680">        
			<td width="178" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atividade</strong></font></td>
			<td width="102" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Interface</strong></font></td>
			<td width="99" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Grupo</strong></font></td>
			<td width="171" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo Processamento </strong></font></td>
			<td width="121" bgcolor="#639ACE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Programa Envolvido</strong></font></td>
		  </tr>	   
	<% end if
	 end if
		if str_Cor = "#EEEEEE" then
			str_Cor = "#FFFFFF"
	   	else
	   		str_Cor = "#EEEEEE"
	   	end if
		
		str_Sql_DadosAdicionais_Tarefa2 = ""
		str_Sql_DadosAdicionais_Tarefa2 = str_Sql_DadosAdicionais_Tarefa & " WHERE PROJ_ID = " & rds_Plano("PLAN_NR_CD_PROJETO_PROJECT")
		str_Sql_DadosAdicionais_Tarefa2 = str_Sql_DadosAdicionais_Tarefa2 & " AND TASK_UID = " & rds_Plano("PLTA_NR_SEQUENCIA_TAREFA") 
		
		'Response.write str_Sql_DadosAdicionais_Tarefa2
		'Response.end 
		
		set rds_DadosAdicionais_Tarefa = db_Cronograma.Execute(str_Sql_DadosAdicionais_Tarefa2)
		if not rds_DadosAdicionais_Tarefa.Eof then
		   dat_Dt_Inicio = rds_DadosAdicionais_Tarefa("TASK_START_DATE")
		   dat_Dt_Termino = rds_DadosAdicionais_Tarefa("TASK_FINISH_DATE")   
		   str_NomeAtividade = rds_DadosAdicionais_Tarefa("TASK_NAME")
		else
		   dat_Dt_Inicio = ""
		   dat_Dt_Termino = ""
		   str_NomeAtividade = ""
		end if
		rds_DadosAdicionais_Tarefa.close
		set rds_DadosAdicionais_Tarefa = Nothing		
	%>
	<% if request("str_Tipo_Saida") <> "Excel" then %>
	 <tr bgcolor="<%=str_Cor%>" width="680">        
        <td colspan="5" bgcolor="<%=str_Cor%>"><img src="../img/001103.gif" width="680" height="1"></td>
      </tr>
	  <% end if %>
      <tr bgcolor="<%=str_Cor%>" width="680">       
        <td width="178" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NomeAtividade%></font></td>
        <td width="102" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Plano("PPAI_TX_CD_INTERFACE")%>-<%=rds_Plano("PPAI_TX_NOME_INTERFACE")%></font></td>
        <td width="99" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Plano("PPAI_TX_GRUPO")%></font></td>
        <td width="171" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Plano("PPAI_TX_TIPO_PROCESSAMENTO")%> </font></td>
        <td width="121" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Plano("PPAI_TX_PROGRAMA_ENVOLVIDO")%></font></td>
      </tr>
    <%  
		rds_Plano.movenext
	Loop 
	rds_Plano.Close
	set rds_Plano = Nothing	
	str_Msg = ""
else
	str_Msg = "Não existem registros para esta condição."
end if	
%>
    </table>
<%
	if str_Msg <> "" then 
	%>
<table width="700"  border="0" cellspacing="0" cellpadding="1">
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