<%
response.buffer=false
%>
<!--#include file="../conn_consulta.asp" -->
<html>
<%
Server.ScriptTimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
set qtos = server.CreateObject("ADODB.RECORDSET")
set qtos_orm = server.CreateObject("ADODB.RECORDSET")

db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

set rs=db.execute("SELECT * FROM SUB_MODULO ORDER BY SUMO_TX_DESC_SUB_MODULO")

set fonte=db.execute("SELECT * FROM ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")
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
<p><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Apoiadores 
  Locais por &Oacute;rg&atilde;o / Processo</strong></font></p>
<table width="52%" border="0" bordercolor="#333333">
  <tr> 
    <td height="42" colspan="2"><img src="topo.gif" width="206" height="32"></td>
    <%
	DO UNTIL RS.EOF=TRUE
	%>
    <td width="37%" align="center" valign="middle" bordercolor="#000000" bgcolor="#CBE1E7"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=RS("SUMO_TX_DESC_SUB_MODULO")%></b></font></div></td>
    <%
	RS.MOVENEXT
	LOOP
	%>
  </tr>
  <%DO UNTIL FONTE.EOF=TRUE%>
  <tr> 
    <td height="26" colspan="2" bgcolor="#CACACA"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=FONTE("AGLU_SG_AGLUTINADO")%> 
      - SEDE</b></font></td>
    <%
	RS.MOVEFIRST
	DO UNTIL RS.EOF=TRUE
	
	SSQL=""
	SSQL="SELECT DISTINCT "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_MATRICULA AS MATRICULA, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR AS LOTACAO, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO AS SITUACAO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR "
	SSQL=SSQL+"FROM         dbo.APOIO_LOCAL_MULT INNER JOIN "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	SSQL=SSQL+"WHERE     (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & FONTE("AGLU_CD_AGLUTINADO") & "') AND "
	SSQL=SSQL+"(dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & RS("SUMO_NR_CD_SEQUENCIA") & ") AND  (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = 1)"
	SSQL=SSQL+"ORDER BY dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR "
	
	SET QTOS=DB.EXECUTE(SSQL)
			
	ATUAL=qtos.recordcount
		
	IF ATUAL=0 THEN
		ATUAL=""
	END IF
	%>
    <td bgcolor="#EEF1EB"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=ATUAL%></strong></font></div></td>
    <%
	RS.MOVENEXT
	LOOP
	%>
  </tr>
  <%
  ORG_AGLU=FORMATNUMBER(FONTE("AGLU_CD_AGLUTINADO"))
  SET ORGAO=DB.EXECUTE("SELECT * FROM ORGAO_MENOR WHERE AGLU_CD_AGLUTINADO=" & ORG_AGLU & " AND ORME_CD_STATUS='A' ORDER BY ORME_SG_ORG_MENOR")%>
  <%
	DO UNTIL ORGAO.EOF=TRUE
	%>
  <tr> 
    <td width="13%" height="26">&nbsp;</td>
    <td width="43%" bgcolor="#E2E2E2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=ORGAO("ORME_SG_ORG_MENOR")%></font></td>
    <%
	RS.MOVEFIRST
	DO UNTIL RS.EOF=TRUE
	
	SSQL=""
	SSQL="SELECT DISTINCT "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_MATRICULA AS MATRICULA, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR AS LOTACAO, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO AS SITUACAO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR "
	SSQL=SSQL+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	SSQL=SSQL+"WHERE     (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & ORGAO("ORME_CD_ORG_MENOR") & "') AND "
	SSQL=SSQL+"(dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & RS("SUMO_NR_CD_SEQUENCIA") & ") AND  (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = 1)"
	SSQL=SSQL+"ORDER BY dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR "
	
	SET QTOS_ORM=DB.EXECUTE(SSQL)
	
	ATUAL_ORM=qtos_orm.recordcount
		
	IF ATUAL_ORM=0 THEN
		ATUAL_ORM=""
	END IF
	
	%>
    <td bgcolor="#CBD3C0"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=ATUAL_ORM%></strong></font></div></td>
    <%
	RS.MOVENEXT
	LOOP
	%>
  </tr>
  <%
	ORGAO.MOVENEXT
	LOOP
	%>
  <%
  FONTE.MOVENEXT
  LOOP
  %>
</table>
<p>&nbsp;</p>
</body>
</html>
