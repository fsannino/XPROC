<!--#include file="conn_consulta.asp" -->
<%
orgao = request("selOrgao")

c1=0
c2=0
c3=0
nc1=0
nc2=0
nc3=0

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

set db2=server.createobject("ADODB.CONNECTION")
db2.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & Server.Mappath("banco.mdb")
db2.CursorLocation = 3

set rs_aglu = db.execute("SELECT * FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO = " & orgao)

ssql=""
ssql="SELECT DISTINCT"
ssql=ssql+" APOIO_LOCAL_MULT.USMA_CD_USUARIO, USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, LEFT(APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR,2) AS ORGAO"
ssql=ssql+" FROM APOIO_LOCAL_MULT"
ssql=ssql+" INNER JOIN USUARIO_MAPEAMENTO ON"
ssql=ssql+" APOIO_LOCAL_MULT.USMA_CD_USUARIO = USUARIO_MAPEAMENTO.USMA_CD_USUARIO"
ssql=ssql+" INNER JOIN APOIO_LOCAL_ORGAO ON"
ssql=ssql+" APOIO_LOCAL_MULT.USMA_CD_USUARIO = APOIO_LOCAL_ORGAO.USMA_CD_USUARIO AND APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO"
ssql=ssql+" WHERE APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=1 AND APOIO_LOCAL_MULT.APLO_NR_SITUACAO=1 AND APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '" & orgao & "%'"
ssql=ssql+" ORDER BY 2,1"

set rs = db.execute(ssql)

Total = rs.recordCount

%>
<html>
<head>
<title>:: Consulta de Capacitação de Apoiadores Locais :::..</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="75%" border="0">
  <tr>
    <td width="84%"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#000099">Consulta 
      de Capacita&ccedil;&atilde;o de Apoiadores Locais</font></b></font></td>
    <td width="8%">&nbsp;</td>
    <td width="8%">&nbsp;</td>
  </tr>
</table>
<p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099">&Oacute;rg&atilde;o 
  Aglutinador Selecionado : <b><%=rs_aglu("AGLU_SG_AGLUTINADO")%></b></font></p>
<table width="94%" border="1" cellpadding="1" cellspacing="0" bordercolor="#E6E6E6">
  <tr bgcolor="#000099"> 
    <td width="10%" height="22"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Chave</font></b></td>
    <td width="37%" height="22"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Nome</font></b></td>
    <td width="21%" height="22"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Assuntos</font></b></td>
    <td width="10%" height="22"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Completeza</font></b></div>
    </td>
    <td width="10%" height="22"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Mapeamento</font></b></div>
    </td>
    <td width="12%" height="22"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Testes 
        Integrados</font></b></div>
    </td>
  </tr>
  <%
  do until rs.eof=true  
  %>
  <tr> 
    <td width="10%" height="21"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("USMA_CD_USUARIO")%></font></td>
    <td width="37%" height="21"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("USMA_TX_NOME_USUARIO")%></font></td>
    <%
		ssql1="SELECT DISTINCT"
		ssql1=ssql1+" APOIO_LOCAL_MODULO.USMA_CD_USUARIO, SUB_MODULO.SUMO_TX_DESC_SUB_MODULO"
		ssql1=ssql1+" FROM APOIO_LOCAL_MODULO"
		ssql1=ssql1+" INNER JOIN SUB_MODULO ON"
		ssql1=ssql1+" APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = SUB_MODULO.SUMO_NR_CD_SEQUENCIA"
		ssql1=ssql1+" WHERE APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO=1 AND APOIO_LOCAL_MODULO.USMA_CD_USUARIO='" & rs("USMA_CD_USUARIO") & "'"
		ssql1=ssql1+" ORDER BY 2"
		
		set temp2=db.execute(ssql1)
		
		assuntos=""
		
		do until temp2.eof=true
			assuntos = assuntos & temp2.fields(1).value & ", "
			temp2.movenext
		loop
		
		if assuntos="" then
			assuntos = "<font color=""white""> - </font>"
		else
			assuntos = left(assuntos, len(assuntos)-2)
		end if
	%>
    <td width="21%" height="21"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=assuntos%></font></td>
    <%
	set temp = db2.execute("SELECT * FROM COMPLETEZA WHERE CHAVE='" & rs("USMA_CD_USUARIO") & "'")
	if temp.eof=true then
		v1="<font color=""white""> - </font>"
		nc1 = nc1 + 1
	else
		v1=" X "
		c1 = c1 + 1
	end if	
	%>
    <td width="10%" height="21"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=v1%></font></b></div>
    </td>
    <%
	set temp = db2.execute("SELECT * FROM MAPEAMENTO WHERE CHAVE='" & rs("USMA_CD_USUARIO") & "'")
	if temp.eof=true then
		v2="<font color=""white""> - </font>"
		nc2 = nc2 + 1
	else
		v2=" X "
		c2 = c2 + 1
	end if	
	%>
    <td width="10%" height="21"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=v2%></font></b></div>
    </td>
    <%
	set temp = db2.execute("SELECT * FROM INTEGRADOS WHERE CHAVE='" & rs("USMA_CD_USUARIO") & "'")
	if temp.eof=true then
		v3="<font color=""white""> - </font>"
		nc3 = nc3 + 1
	else
		v3=" X "
		c3 = c3 + 1
	end if	
	%>
    <td width="12%" height="21"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=v3%></font></b></div>
    </td>
  </tr>
  <%
  rs.movenext
  loop
  %>
  <tr> 
    <td width="10%" height="46"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
    <td width="37%" height="46"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
    <td width="21%" height="46"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
    <td width="10%" height="46"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
    <td width="10%" height="46"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
    <td width="12%" height="46"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
  </tr>
  <tr> 
    <td width="10%" height="26" bordercolor="#E2E2E2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
    <td width="37%" bgcolor="#FFFFFF" bordercolor="#E2E2E2" height="26"><b></b></td>
    <td width="21%" bordercolor="#999999" height="26" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Total 
      de Apoiadores Locais</font></b></td>
    <td width="10%" bordercolor="#999999" height="26" bgcolor="#FFFFFF"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=TOTAL%></font></b></div>
    </td>
    <td width="10%" bordercolor="#999999" height="26" bgcolor="#FFFFFF"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=TOTAL%></font></b></div>
    </td>
    <td width="12%" bordercolor="#999999" height="26" bgcolor="#FFFFFF"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=TOTAL%></font></b></div>
    </td>
  </tr>
  <tr> 
    <td width="10%" height="24" bordercolor="#E2E2E2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
    <td width="37%" bgcolor="#FFFFFF" bordercolor="#E2E2E2" height="24"><b></b></td>
    <td width="21%" bordercolor="#999999" height="24" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Capacitados</font></b></td>
    <td width="10%" bordercolor="#999999" height="24" bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=c1%></font></div>
    </td>
    <td width="10%" bordercolor="#999999" height="24" bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=c2%></font></div>
    </td>
    <td width="12%" bordercolor="#999999" height="24" bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=c3%></font></div>
    </td>
  </tr>
  <tr> 
    <td width="10%" height="25" bordercolor="#E2E2E2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></td>
    <td width="37%" bgcolor="#FFFFFF" bordercolor="#E2E2E2" height="25"><b></b></td>
    <td width="21%" bordercolor="#999999" height="25" bgcolor="#CCCCCC"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">N&atilde;o 
      Capacitados</font></b></td>
    <td width="10%" bordercolor="#999999" height="25" bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=nc1%></font></div>
    </td>
    <td width="10%" bordercolor="#999999" height="25" bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=nc2%></font></div>
    </td>
    <td width="12%" bordercolor="#999999" height="25" bgcolor="#FFFFFF"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=nc3%></font></div>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
