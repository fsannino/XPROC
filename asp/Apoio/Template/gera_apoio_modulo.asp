<!--#include file="../conn_consulta.asp" -->
<%
'Response.Buffer = False

if request("excel")=1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
<!--#include file="../conn_consulta.asp" -->
<html>
<%
server.ScriptTimeout=99999999

tem=0
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

db.cursorlocation=3

aglu=request("selAglu")

set rs=db.execute("SELECT AGLU_SG_AGLUTINADO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & aglu)
txtaglu=rs("AGLU_SG_AGLUTINADO")

set fonte = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM APOIO_LOCAL_ORGAO WHERE ORME_CD_ORG_MENOR LIKE '" & aglu & "%' AND APLO_NR_ATRIBUICAO=" & request("selAtrib") & " ORDER BY ORME_CD_ORG_MENOR")

%>
<head>
<title>Base de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>
<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
if request("excel")=0 then
%>
<table width="67%" height="26" border="0">
  <tr> 
    <td width="5%"><div align="right"><a href="javascript:history.go(-1)"><img src="../volta_f02.gif" width="24" height="24" border="0"></a></div></td>
    <td width="15%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Voltar</font></strong></td>
    <td width="5%"><div align="right"><a href="javascript:print()"><img src="../impress%E3o.jpg" width="27" height="21" border="0"></a></div></td>
    <td width="16%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Imprimir</font></strong></td>
    <td width="6%"><div align="right"><a href="gera_apoio_modulo.asp?excel=1&selAglu=<%=request("selAglu")%>&selClass=<%=request("selClass")%>&selAtrib=<%=request("selAtrib")%>" target="_blank"><img src="../excel.jpg" width="22" height="20" border="0"></a></div></td>
    <td width="53%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Exportar 
      para o Excel</font></strong></td>
  </tr>
</table>
<%end if%>
<p><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Oacute;rg&atilde;o 
  Aglutinador Selecionado : <%=txtaglu%></strong></font></p>
<p>
  <%
if request("selClass")=1 then
%>
</p>
<%
ssql=""
ssql="SELECT DISTINCT Left([ORME_CD_ORG_MENOR],2) AS ORGAO, APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, APOIO_LOCAL_ORGAO.USMA_CD_USUARIO"
ssql=ssql+" FROM APOIO_LOCAL_ORGAO "
ssql=ssql+"WHERE (((Left([ORME_CD_ORG_MENOR],2))=" & aglu & ") AND ((APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO)=" & request("selAtrib") & "))"

set quantos=db.execute(ssql)

qtos=quantos.recordcount

%>
<%if request("selAtrib")=1 then%>
<p><font color="#003366" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#000066" size="2">Total 
  de Apoiadores Locais Encontrados : </font></strong><font color="#000066" size="2"><%=qtos%></font></font></p>
<font color="#000066" size="2"><strong><font face="Verdana, Arial, Helvetica, sans-serif"> 
<%else%>
</font></strong></font><font face="Verdana, Arial, Helvetica, sans-serif">
<p><font color="#000066" size="2"><strong>Total de Multiplicadores Encontrados : 
  </strong><%=qtos%><strong> </strong></font></p>
</font><%end if%>
<table width="85%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="28%"><strong><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;RG&Atilde;O 
      APOIADO</font></strong></td>
    <td colspan="2"><strong><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif">ASSUNTO</font></strong> 
      <div align="center"></div></td>
  </tr>
  <%
do until fonte.eof=true
%>
  <tr> 
    <%
	if len(fonte("ORME_CD_ORG_MENOR"))=2 then
		set orgao=db.execute("SELECT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & fonte("ORME_CD_ORG_MENOR"))
		pre=" - SEDE"
	else
		set orgao=db.execute("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & fonte("ORME_CD_ORG_MENOR") & "'")	
		pre=""
	end if	
	
	org_=orgao("ORGAO")
	%>
    <td bgcolor="#D1D1D1"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=org_%></font></strong></td>
    <td width="66%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
    <td width="6%"><div align="center"></div></td>
  </tr>
  <%
ssql=""
ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      COUNT(DISTINCT dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO) AS CONTA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
ssql=ssql+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "

ssql=ssql+"HAVING      (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND "

ssql=ssql+"                      (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("ORME_CD_ORG_MENOR") & "')"

set rs_org=db.execute(ssql)

CONTA=0

do until rs_org.eof=true
%>
  <tr> 
    <td><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
    <td bgcolor="#E6E6E6"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_org("SUMO_TX_DESC_SUB_MODULO")%></font></strong></td>
    <td><div align="center"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_org("CONTA")%></font></div></td>
  </tr>
  <%
  CONTA=CONTA+rs_org("CONTA")
  tem=tem+1
rs_org.movenext
loop
%>
  <tr> 
    <td><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td><div align="right"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
        por &Oacute;rg&atilde;o Apoiado : </strong></font></div></td>
    <td><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><B><%=CONTA%></B></font></div></td>
  </tr>
  <%
fonte.movenext
loop
%>
</table>
<%else%>
<p>&nbsp;</p>
<table width="93%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="23%"><strong><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;RG&Atilde;O 
      APOIADO</font></strong></td>
    <td colspan="5"><strong><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif">ASSUNTO</font></strong>
<div align="center"></div></td>
  </tr>
  <%
do until fonte.eof=true
%>
  <tr> 
    <%
	if len(fonte("ORME_CD_ORG_MENOR"))=2 then
		set orgao=db.execute("SELECT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & fonte("ORME_CD_ORG_MENOR"))
	else
		set orgao=db.execute("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & fonte("ORME_CD_ORG_MENOR") & "'")	
	end if	
	
	org_=orgao("ORGAO")
	%>
    <td bgcolor="#D1D1D1"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=org_%></font></strong></td>
    <td colspan="4"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
    <td width="8%"><div align="center"></div></td>
  </tr>
  <%
ssql=""
ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      COUNT(DISTINCT dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO) AS CONTA, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
ssql=ssql+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "

ssql=ssql+"HAVING      (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND "

ssql=ssql+"                      (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("ORME_CD_ORG_MENOR") & "')"

set rs_org=db.execute(ssql)

CONTA=0

do until rs_org.eof=true
%>
  <tr> 
    <td><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></strong></td>
    <td colspan="4" bgcolor="#E6E6E6"><strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs_org("SUMO_TX_DESC_SUB_MODULO")%></font></strong></td>
    <td><div align="center"></div></td>
    <%
	CONTADOR = rs_org("CONTA")
	CONTA=CONTA+rs_org("CONTA")
	%>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="4">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td width="8%" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="32%" bgcolor="#DADADA"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></td>
    <td width="21%" bgcolor="#DADADA"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Lota&ccedil;&atilde;o</strong></font></td>
    <td width="8%" bgcolor="#DADADA"><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Chave</strong></font></td>
    <td>&nbsp;</td>
  </tr>
  <%
ssql=""
ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO AS CHAVE, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR AS LOTACAO "
ssql=ssql+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
ssql=ssql+"                      dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"                      dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, "
ssql=ssql+"                      dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO, dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL, "
ssql=ssql+"                      dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR "
ssql=ssql+"HAVING      (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & request("selAtrib") & ") AND "
ssql=ssql+"                      (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("ORME_CD_ORG_MENOR") & "') AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs_org("SUMO_NR_CD_SEQUENCIA") & ") ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"

set nomes=db.execute(ssql)

do until nomes.eof=true
%>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=nomes("USMA_TX_NOME_USUARIO")%></font></td>
	<%
	SET TEMP=DB.EXECUTE("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & nomes("LOTACAO") & "'")
	LOT=TEMP("ORGAO")
	%>
    <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=LOT%></font></td>
    <td><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=nomes("CHAVE")%></font></td>
    <td>&nbsp;</td>
  </tr>
  <%
    tem=tem+1
nomes.movenext
loop
%>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="4" bgcolor="#FFFFFF"> 
      <div align="right"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
        por Assunto: </strong></font></div></td>
    <td bgcolor="#FFFFFF"> 
      <div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=CONTADOR%></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="4">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
<%
rs_org.movenext
loop
%>
  <tr> 
    <td>&nbsp;</td>
    <td colspan="4">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td colspan="4"><div align="right"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
        por &Oacute;rg&atilde;o Apoiado : </strong></font></div></td>
    <td><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><B><%=CONTA%></B></font></div></td>
  </tr>
  <%
response.Flush()
fonte.movenext
loop
%>
</table>
<%end if%>
<%if tem=0 then%>
<p><font color="#800000"><strong>Nenhum Registro Encontrado para a Sele&ccedil;&atilde;o</strong></font></p>
<%end if%>
</body>
</html>
