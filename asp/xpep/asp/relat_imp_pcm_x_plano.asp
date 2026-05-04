<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Expires=0

IF request("str_Tipo_Saida")="Excel" THEN
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
END IF

int_Onda = request("selOnda")
int_Fase = request("selFases")
int_Plano1 = request("selPlano")

int_Plano2 = request("selPlano2")
int_Atividade = request("selTask1")

'response.Write(int_Onda)
'response.Write(int_Fase)
'response.Write(int_Plano1)
'response.Write(int_Plano2)
'response.Write(int_Atividade)
'response.end 

if int_Plano1 <> "" then	
	vet_int_Plano1 = Split(int_Plano1,"|")
	int_Plano1 = vet_int_Plano1(0)
end if
if int_Plano2 <> "" then	
	vet_int_Plano2 = Split(int_Plano2,"|")
	int_Plano2 = vet_int_Plano2(0)
end if

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")
db_Cogest.cursorlocation = 3

str_SQL = ""
str_SQL = str_SQL & " SELECT distinct "
str_SQL = str_SQL & " dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_TX_SIGLA_PLANO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_TX_DESCRICAO_PLANO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_FASE"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_ONDA"
str_SQL = str_SQL & " , dbo.ONDA.ONDA_TX_DESC_ONDA"
str_SQL = str_SQL & " FROM dbo.ONDA INNER JOIN"
str_SQL = str_SQL & " dbo.XPEP_PLANO_ENT_PRODUCAO ON dbo.ONDA.ONDA_CD_ONDA = dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_ONDA INNER JOIN"
str_SQL = str_SQL & " dbo.XPEP_PLANO_TAREFA_PCM ON "
str_SQL = str_SQL & " dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO = dbo.XPEP_PLANO_TAREFA_PCM.PLAN_NR_SEQUENCIA_PLANO"
str_SQL = str_SQL & " WHERE dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO > 0 "
if int_Onda <> "" then
	str_SQL = str_SQL & " AND dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_ONDA = " & int_Onda
end if
if int_Fase <> "" then
	str_SQL = str_SQL & " AND dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_FASE = " & int_Fase
end if
if int_Plano2 <> "" then
	str_SQL = str_SQL & " AND dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO = " & int_Plano2
end if
str_SQL = str_SQL & " ORDER BY "
str_SQL = str_SQL & " dbo.ONDA.ONDA_TX_DESC_ONDA"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_CD_FASE"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_TX_SIGLA_PLANO"
'response.Write(str_SQL)
'response.End()
set rds_Plano = db_Cogest.Execute(str_SQL)

str_SQL = ""
str_SQL = str_SQL & " SELECT "
str_SQL = str_SQL & " PLAN_NR_SEQUENCIA_PLANO"
str_SQL = str_SQL & " , PPCM_NR_SEQUENCIA_TAREFA"
str_SQL = str_SQL & " , PPCM_TX_ATIVIDADE"
str_SQL = str_SQL & " , PPCM_TX_TP_COMUNICACAO "
str_SQL = str_SQL & " , PPCM_TX_O_QUE_COMUNICAR"
str_SQL = str_SQL & " , PPCM_TX_PARA_QUEM"
str_SQL = str_SQL & " , PPCM_TX_UNID_ORGAO"
str_SQL = str_SQL & " , PPCM_TX_QUANDO_OCORRE"
str_SQL = str_SQL & " , PPCM_TX_RESP_CONTEUDO "
str_SQL = str_SQL & " , PPCM_TX_RESP_DIVULGACAO"
str_SQL = str_SQL & " , PPCM_TX_COMO"
str_SQL = str_SQL & " , PPCM_TX_APROVADOR_PB"
str_SQL = str_SQL & " , PPCM_DT_APROVACAO"
str_SQL = str_SQL & " , PPCM_TX_ARQUIVO_ANEXO1 "
str_SQL = str_SQL & " , PPCM_TX_ARQUIVO_ANEXO2"
str_SQL = str_SQL & " , PPCM_TX_ARQUIVO_ANEXO3"
str_SQL = str_SQL & " FROM  dbo.XPEP_PLANO_TAREFA_PCM"

Function FormataData(str_Data)

	strDia = ""		
	strMes = ""
	strAno = ""
	vet_Data = split(Trim(str_Data),"/")							
	strDia = trim(vet_Data(1))
	if cint(strDia) < 10 then
		strDia = "" & strDia
	end if			
	strMes = trim(vet_Data(0))			
	if cint(strMes) < 10 then
		strMes = "" & strMes
	end if
	strAno = trim(vet_Data(2))
	str_data = strDia & "/" & strMes & "/" & strAno
	FormataData = str_data
	
end function

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
    <table width="660"  border="0" cellspacing="0" cellpadding="1">
      <tr>
        <td width="991"><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Relat&oacute;rio do Plano de Comunica&ccedil;&atilde;o - PCM </font></strong></div></td>
      </tr>
      <tr>
        <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ordenado por onda - fase - plano e data limite para comunica&ccedil;&atilde;o</font></div></td>
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
		<table width="660"  border="0" cellspacing="5" cellpadding="1">
		  <tr>
			<td width="194">				  
			  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
			  <strong>
			  <% if boo_MostraOnda then %>
			  Onda -<%=rds_Plano("ONDA_TX_DESC_ONDA")%>
		      <% end if %>
	        </strong>			    </font></td>
			<td width="160">
			  <strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
			  <% if boo_MostraFase then %>
			  Fase -<%=rds_Plano("PLAN_NR_CD_FASE")%>
			  <% end if %>
	        </font></strong></td>
			<td width="74"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Plano - <%=rds_Plano("PLAN_TX_SIGLA_PLANO")%></font></strong></td>
			<td width="199"><div align="right"></div></td>
		  </tr>
</table>
<%	
		end if
	int_Cont = 0
	str_SQL2 = ""
	str_SQL2 = str_SQL2 & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & rds_Plano("PLAN_NR_SEQUENCIA_PLANO")
    str_SQL2 = str_SQL2 & " ORDER BY PPCM_TX_QUANDO_OCORRE "
    set rds_PCM = db_Cogest.Execute(str_SQL + str_SQL2)
	 %>
<table width="680"  border="0" cellpadding="0" cellspacing="3" bordercolor="#CCCCCC">
<% if boo_Mostra_Cabec = True then %>	
<% if request("str_Tipo_Saida") <> "Excel" then %>
  <tr bgcolor="<%=str_Cor%>">
    <td colspan="6" bgcolor="<%=str_Cor%>"><img src="../img/tit_tab_imp_Fundo_PCM1.gif" width="680" height="25"></td>
  </tr>
<% else %>  
  <tr bgcolor="#639ACE">
    <td><div align="center"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atividade</font></strong></div></td>
    <td><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data limite para comuinica&ccedil;&atilde;o</font></strong></td>
    <td><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Para quem comunicar</font></strong></td>
    <td><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">O que comunicar </font></strong></td>
    <td><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Tipo comunica&ccedil;&atilde;o</font></strong></td>
    <td><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Unidade/&Oacute;rg&atilde;o</font></strong></td>
  </tr>
<% end if %>  
<% end if %>
  <%
	int_TotRegistroPCM = rds_PCM.recordcount
	int_LoopPCM = 0
	do until int_TotRegistroPCM = int_LoopPCM
	   int_LoopPCM = int_LoopPCM + 1
	%>
  <% 	if str_Cor = "#EEEEEE" then
			str_Cor = "#FFFFFF"
	   	else
	   		str_Cor = "#EEEEEE"
	   	end if
		if Trim(rds_PCM("PPCM_TX_TP_COMUNICACAO")) = "INT" then
			str_Tp_Comunic = "INTERNO"
		elseif Trim(rds_PCM("PPCM_TX_TP_COMUNICACAO")) = "EXT" then
			str_Tp_Comunic = "EXTERNO"
		else
			str_Tp_Comunic = "INTERNO/ EXTERNO"
		end if
	%>
	<% if request("str_Tipo_Saida") <> "Excel" then %>
  <tr bgcolor="<%=str_Cor%>">
    <td colspan="6" bgcolor="<%=str_Cor%>"><img src="../img/001103.gif" width="680" height="1"></td>
  </tr>
  	<% end if %>
  <tr bgcolor="<%=str_Cor%>">
    <td width="96" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCM("PPCM_TX_ATIVIDADE")%></font></td>
    <td width="117" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormataData(rds_PCM("PPCM_TX_QUANDO_OCORRE"))%></font></td>
    <td width="89" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCM("PPCM_TX_PARA_QUEM")%></font></td>
    <td width="138" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCM("PPCM_TX_O_QUE_COMUNICAR")%></font></td>
    <td width="133" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Tp_Comunic%> </font></td>
    <td width="92" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_PCM("PPCM_TX_UNID_ORGAO")%></font></td>
  </tr>
  <%  
			rds_PCM.movenext
		Loop
		rds_PCM.close
		rds_Plano.movenext
	Loop 
	rds_Plano.Close
	set rds_Plano = Nothing
	set rds_PCM = Nothing
	str_Msg = ""
else
	str_Msg = "Não existem registros para esta condição."
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
