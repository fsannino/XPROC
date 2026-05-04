<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Expires=0

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
'str_SQL = str_SQL & " SELECT distinct "
str_SQL = str_SQL & " SELECT "
str_SQL = str_SQL & " dbo.XPEP_PLANO_TAREFA_PCM.PLAN_NR_SEQUENCIA_PLANO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_NR_SEQUENCIA_TAREFA"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_ATIVIDADE"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_TP_COMUNICACAO "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_O_QUE_COMUNICAR"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_PARA_QUEM"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_UNID_ORGAO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_QUANDO_OCORRE"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_RESP_CONTEUDO "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_RESP_DIVULGACAO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_COMO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_APROVADOR_PB"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_DT_APROVACAO"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_ARQUIVO_ANEXO1 "
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_ARQUIVO_ANEXO2"
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_ARQUIVO_ANEXO3"

str_SQL = str_SQL & " , dbo.XPEP_PLANO_ENT_PRODUCAO.PLAN_NR_SEQUENCIA_PLANO"
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
str_SQL = str_SQL & " , dbo.XPEP_PLANO_TAREFA_PCM.PPCM_TX_QUANDO_OCORRE"
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
		strDia = "0" & CInt(strDia)
	end if			
	strMes = trim(vet_Data(0))			
	if cint(strMes) < 10 then
		strMes = "0" & CInt(strMes)
	end if
	strAno = trim(vet_Data(2))
	str_data = strDia & "/" & strMes & "/" & strAno
	FormataData = str_data
	
end function

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

function impressao() 
	{
	// selOnda=<%=int_Onda%>&selFases=<%=int_Fase%>&selPlano=<%=int_Plano1%>&selPlano2=<%=int_Plano2%>&selTask1=<%=int_Atividade%>
	window.open('impressao.asp?par_PaginaPrint=relat_imp_pcm_x_plano_data.asp&selOnda=<%=int_Onda%>&selFases=<%=int_Fase%>&selPlano=<%=int_Plano1%>&selPlano2=<%=int_Plano2%>&selTask1=<%=int_Atividade%>','jan1','toolbar=no, location=no, scrollbars=no, status=no, directories=no, resizable=no, menubar=no, fullscreen=no, height=50, width=250, status=no, top=200, left=260');
	}
	
function exporta() 
	{
	// selOnda=<%=int_Onda%>&selFases=<%=int_Fase%>&selPlano=<%=int_Plano1%>&selPlano2=<%=int_Plano2%>&selTask1=<%=int_Atividade%>
	window.open('relat_imp_pcm_x_plano_data.asp?str_Tipo_Saida=Excel&selOnda=<%=int_Onda%>&selFases=<%=int_Fase%>&selPlano=<%=int_Plano1%>&selPlano2=<%=int_Plano2%>&selTask1=<%=int_Atividade%>','jan1','toolbar=yes, location=no, scrollbars=yes, status=no, directories=no, resizable=yes, menubar=yes, fullscreen=no, height=400, width=500, status=no, top=100, left=160');
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
	<!-- InstanceBeginEditable name="corpo" -->    <table width="77%"  border="0" cellspacing="0" cellpadding="1">
      <tr>
        <td width="10%">&nbsp;</td>
        <td width="76%">&nbsp;</td>
        <td width="14%">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td class="subtitulob"><div align="center" class="campob">Relat&oacute;rio do Plano de Comunica&ccedil;&atilde;o - PCM </div></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ordenado por onda - fase e data limite para comunica&ccedil;&atilde;o</font></div></td>
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
		boo_Mostra_Cabec = False
		int_LoopPlano = int_LoopPlano + 1
		str_Cor = "#FFFFFF"
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
                <td width="204"><div align="right">
                    <div align="left"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
                      <% if boo_MostraOnda then %>
                      </strong>Onda -<%=rds_Plano("ONDA_TX_DESC_ONDA")%><strong>
                      <% end if %>
                    </strong></font></strong></div>
                </div></td>
                <td width="254" class="campob"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                  <% if boo_MostraFase then %>
      Fase - <%=rds_Plano("PLAN_NR_CD_FASE")%>
      <% end if %>
                </font></strong></td>
                <td width="94">&nbsp;</td>
                <td width="85">&nbsp;</td>
              </tr>
            </table>
			<% end if %>
	<% 	
	int_Cont = 0
	 %>
	<table width="800"  border="0" cellpadding="1" cellspacing="2" bordercolor="#CCCCCC">
	<% if boo_Mostra_Cabec = True then %>	
      <tr width="800">
        <td width="1" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="32" bgcolor="#6396CE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Plano</strong></font></td>
        <td width="157" bgcolor="#6396CE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Atividade</strong></font></td>
        <td width="102" bgcolor="#6396CE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data Limite para Comumica&ccedil;&atilde;o</strong></font></td>
        <td width="113" bgcolor="#6396CE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Pra Quem Comunicar</strong></font></td>
        <td width="133" bgcolor="#6396CE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>O que comunicar</strong></font></td>
        <td width="136" bgcolor="#6396CE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo Comunica&ccedil;&atilde;o</strong></font></td>
        <td width="103" bgcolor="#6396CE"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><span class="style23">Unidade/&Oacute;rg&atilde;o</span></strong></font></td>
      </tr>	
	<% end if
		if str_Cor = "#EEEEEE" then
			str_Cor = "#FFFFFF"
	   	else
	   		str_Cor = "#EEEEEE"
	   	end if
		if Trim(rds_Plano("PPCM_TX_TP_COMUNICACAO")) = "INT" then
			str_Tp_Comunic = "INTERNO"
		elseif Trim(rds_Plano("PPCM_TX_TP_COMUNICACAO")) = "EXT" then
			str_Tp_Comunic = "EXTERNO"
		else
			str_Tp_Comunic = "INTERNO/EXTERNO"
		end if
	%>
	 <tr bgcolor="<%=str_Cor%>">
        <td bgcolor="#FFFFFF"></td>
        <td colspan="8" bgcolor="<%=str_Cor%>"><img src="../img/001103.gif" width="800" height="1"></td>
      </tr>
	  
      <tr bgcolor="<%=str_Cor%>" width="800">
        <td width="1" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="32" bgcolor="<%=str_Cor%>" align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=rds_Plano("PLAN_TX_SIGLA_PLANO")%></b></font></td>
        <td width="157" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="inclui_altera_plano_pcm_sub.asp?pAcao=C&pPlano=<%=rds_Plano("PLAN_NR_SEQUENCIA_PLANO")%>&pPlano2=<%=rds_Plano("PLAN_NR_SEQUENCIA_PLANO")%>&pOnda=<%=rds_Plano("PLAN_NR_CD_ONDA")%>&pCdSeqPCM=<%=rds_Plano("PPCM_NR_SEQUENCIA_TAREFA")%>&pPlano_Origem=<%=rds_Plano("PLAN_TX_SIGLA_PLANO")%>" class="link"><%=rds_Plano("PPCM_TX_ATIVIDADE")%></a></font></td>
        <td width="102" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FormataData(rds_Plano("PPCM_TX_QUANDO_OCORRE"))%></font></td>
        <td width="113" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Plano("PPCM_TX_PARA_QUEM")%></font></td>
        <td width="133" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Plano("PPCM_TX_O_QUE_COMUNICAR")%></font></td>
        <td width="136" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Tp_Comunic%> </font></td>
        <td width="103" bgcolor="<%=str_Cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Plano("PPCM_TX_UNID_ORGAO")%></font></td>
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
        <td><a href="#"><img src="../img/botao_imprimir.gif" alt=":: Imprime formato relat&oacute;rio" width="85" height="19" border="0" onclick="impressao();"></a></td>
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
