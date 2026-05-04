<!--#include file="../conn_consulta.asp" -->
<html>
<%
tem=0
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

db.cursorlocation=3

set rs=db.execute("SELECT AGLU_CD_AGLUTINADO, AGLU_SG_AGLUTINADO FROM ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")
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

<table width="67%" height="26" border="0">
  <tr> 
    <td width="5%"><div align="right"><a href="javascript:history.go(-1)"><img src="../volta_f02.gif" width="24" height="24" border="0"></a></div></td>
    <td width="15%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Voltar</font></strong></td>
    <td width="5%"><div align="right"><a href="javascript:print()"><img src="../impress%E3o.jpg" width="27" height="21" border="0"></a></div></td>
    <td width="16%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Imprimir</font></strong></td>
    <td width="6%"><div align="right"></div></td>
    <td width="53%"><strong></strong></td>
  </tr>
</table>
<p><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
  Geral de Apoiadores Locais / Multiplicadores por &Oacute;rg&atilde;o Apoiado</strong></font> 
</p>
<table width="60%" border="1" bordercolor="#000000">
  <tr bgcolor="#CCCCCC"> 
    <td width="47%"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Org&atilde;o 
      Aglutinador </strong></font></td>
    <td width="28%"><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Apoiadores 
        Locais </strong></font></div></td>
    <td width="25%"><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Multiplicadores</strong></font></div></td>
  </tr>
  <%
  VALOR_1=1
  VALOR_2=2
  
	i=0
	reg=rs.RecordCount	
  
  do until i = reg%>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><%=RS("AGLU_SG_AGLUTINADO")%></b></font>&nbsp;</td>
    <%
    ssql=""
    ssql="SELECT DISTINCT dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"	FROM dbo.APOIO_LOCAL_MULT "
	ssql=ssql+"INNER JOIN dbo.APOIO_LOCAL_ORGAO ON "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"WHERE  (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO=1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=" & VALOR_1 & ") AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '" & RS("AGLU_CD_AGLUTINADO") & "%') "
	
	set quantos1=db.execute(ssql)

	qtos1 = quantos1.recordcount
	
	total_apoio=total_apoio+qtos1
	%>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=qtos1%></font></div></td>
    <%
    ssql=""
    ssql="SELECT DISTINCT dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"	FROM dbo.APOIO_LOCAL_MULT "
	ssql=ssql+"INNER JOIN dbo.APOIO_LOCAL_ORGAO ON "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"WHERE  (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO=1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO=" & VALOR_2 & ") AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '" & RS("AGLU_CD_AGLUTINADO") & "%') "
	
	set quantos2=db.execute(ssql)

	qtos2 = quantos2.recordcount
	
	total_mult=total_mult+qtos2
	%>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=qtos2%></font></div></td>
  </tr>
  <%
  i=i+1
  rs.movenext
  loop  
  %>
  <tr> 
    <td height="33" bgcolor="#CCCCCC"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Total 
        Geral =&gt;</b></font></div></td>
    <td><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=total_apoio%></font></strong></div></td>
    <td><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=total_mult%></font></strong></div></td>
  </tr>
</table>
<p>&nbsp;</p></body>
</html>