<%
response.buffer=false
%>
<!--#include file="../conn_consulta.asp" -->
<html>
<%
Server.ScriptTimeout=99999999

set db=server.createobject("ADODB.CONNECTION")

db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

set fonte=db.execute("SELECT * FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & request("selAglu") & " ORDER BY AGLU_SG_AGLUTINADO")

txtaglu=fonte("AGLU_SG_AGLUTINADO")

set rs=db.execute("SELECT * FROM SUB_MODULO ORDER BY SUMO_TX_DESC_SUB_MODULO")
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
    <td width="6%"><div align="right"><img src="../excel.jpg" width="22" height="20" border="0"></div></td>
    <td width="53%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Exportar 
      para o Excel</font></strong></td>
  </tr>
</table>
<p><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Relat&oacute;rio 
  de M&oacute;dulos n&atilde;o Apoiados por &Oacute;rg&atilde;o</strong></font></p>
<table width="77%" border="0">
  <tr> 
    <td width="59%" height="63" rowspan="2"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Oacute;rg&atilde;o 
      Aglutinador Selecionado : <%=txtaglu%></strong></font></td>
    <td width="6%"> <div align="center"></div></td>
    <td width="35%"><font color="BLUE" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong>X</strong></font><font color="#FF0000" size="4" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
      - <strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;o 
      apoiado </font></strong></td>
  </tr>
  <tr> 
    <td><div align="center"></div></td>
    <td><font color="#FF0000" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong>X 
      </strong></font>- <strong><font color="#000066" size="1" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;o 
      n&atilde;o apoiado</font></strong></td>
  </tr>
</table>
<table width="73%" border="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="77%" border="0">
  <tr bgcolor="#000066"> 
    <td height="40" colspan="3" bgcolor="#FFFFFF"> 
      <div align="center"><font color="#FFFFFF" size="1"><strong><img src="topo.gif" width="206" height="32"></strong></font></div></td>
    <%
	rs.movefirst
	do until rs.eof=true
	%>
    <td width="33%"><div align="center"> 
        <h6><font color="#FFFFFF" size="1"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUMO_TX_DESC_SUB_MODULO")%></font></strong></font></h6>
      </div></td>
    <%
	rs.movenext
	loop
	%>
  </tr>
  <%
  do until fonte.eof=true
  %>
  <tr> 
    <td colspan="3" bgcolor="<%=cor%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><B><%=fonte("AGLU_SG_AGLUTINADO")%></B></font></td>
    <%
	rs.movefirst
	do until rs.eof=true

	ssql=""
	ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO AS Expr1, "
	ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO AS Expr2, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, "
	ssql=ssql+"	                      dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR "
	ssql=ssql+"	FROM         dbo.APOIO_LOCAL_MULT INNER JOIN"
	ssql=ssql+"	                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN"
	ssql=ssql+"	                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"	WHERE     (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = 1) AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = 1) AND "
	ssql=ssql+"	                  (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND "
	ssql=ssql+"                      (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs("SUMO_NR_CD_SEQUENCIA") & ") AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & fonte("AGLU_CD_AGLUTINADO") & "')"
	
	set org1=db.execute(ssql)
	
	MARCA="X"
	
	if org1.eof=true then
	%>
    <td width="13%" bgcolor="#D8E7E7"><div align="center"><font color="#FF0000" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=marca%></strong></font></div></td>
	<%
	else
	%>
    <td width="13%" bgcolor="#D8E7E7"><div align="center"><font color="BLUE" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=marca%></strong></font></div></td>	
	<%
	end if
	rs.movenext
	loop
	%>
  </tr>
  <%
  set maior=db.execute("SELECT * FROM ORGAO_MAIOR WHERE AGLU_CD_AGLUTINADO=" & fonte("AGLU_CD_AGLUTINADO") & " AND ORLO_CD_STATUS='A' ORDER BY ORLO_SG_ORG_LOT")
  do until maior.eof=true
  
	aglutinador = maior("AGLU_CD_AGLUTINADO")
    lotacao = maior("ORLO_CD_ORG_LOT")
    sequencia = maior("ORLO_NR_ORDEM")
    
    prefixo_orgao = Right("000" & aglutinador, 2) & Right("0000" & lotacao, 3) & (Right("000" & sequencia, 2))
    prefixo_orgao = Trim(prefixo_orgao)

  %>
  <tr> 
    <td width="8%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td colspan="2" bgcolor="#E2E2E2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><B><%=maior("ORLO_SG_ORG_LOT")%></B></font></td>
    <%
	rs.movefirst
	do until rs.eof=true

	ssql=""
	ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO AS Expr1, "
	ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO AS Expr2, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, "
	ssql=ssql+"	                      dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR "
	ssql=ssql+"	FROM         dbo.APOIO_LOCAL_MULT INNER JOIN"
	ssql=ssql+"	                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN"
	ssql=ssql+"	                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"	WHERE     (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = 1) AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = 1) AND "
	ssql=ssql+"	                  (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND "
	ssql=ssql+"                      (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs("SUMO_NR_CD_SEQUENCIA")& ") AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & prefixo_orgao & "00000000')"
	
	set org2=db.execute(ssql)
	
	MARCA="X"

	if org2.eof=true then
	%>
	<td bgcolor="#D8E7E7"><div align="center"><font color="#FF0000" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=marca%></strong></font></div></td>
	<%
	ELSE
	%>
	<td bgcolor="#D8E7E7"><div align="center"><font color="BLUE" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=marca%></strong></font></div></td>	
	<%
	END IF
	rs.movenext
	loop
	%>
  </tr>
  <%
	set menor = db.execute("SELECT * FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR LIKE '" & prefixo_orgao & "%' ORDER BY ORME_SG_ORG_MENOR")
	do until menor.eof=true
	if trim(menor("ORME_CD_ORG_MENOR"))<>trim(prefixo_orgao & "00000000") then
  %>
  <tr> 
    <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="8%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td width="38%" bgcolor="#F0F0F0"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=menor("ORME_SG_ORG_MENOR")%></font></td>
    <%
	rs.movefirst
	do until rs.eof=true

	ssql=""
	ssql="SELECT     dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO AS Expr1, "
	ssql=ssql+"                      dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO AS Expr2, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, "
	ssql=ssql+"	                      dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR "
	ssql=ssql+"	FROM         dbo.APOIO_LOCAL_MULT INNER JOIN"
	ssql=ssql+"	                      dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN"
	ssql=ssql+"	                      dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"	WHERE     (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = 1) AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = 1) AND "
	ssql=ssql+"	                  (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = 1) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND "
	ssql=ssql+"                      (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs("SUMO_NR_CD_SEQUENCIA") & ") AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & menor("ORME_CD_ORG_MENOR") & "')"
	
	set org3=db.execute(ssql)
	
	MARCA="X"
	
	if org3.eof=true then
	%>
	<td bgcolor="#D8E7E7"><div align="center"><font color="#FF0000" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong>X</strong></font></div></td>
	<%
	ELSE
	%>
	<td bgcolor="#D8E7E7"><div align="center"><font color="BLUE" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong>X</strong></font></div></td>
	<%
	END IF
	rs.movenext
	loop
	%>
  </tr>
  <%
	end if
  	menor.movenext
	loop
    maior.movenext
	loop
	fonte.movenext
	loop  
  %>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></strong></p>
<p>&nbsp;</p>
</body>
</html>
