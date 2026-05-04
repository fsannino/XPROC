<!--#include file="../conn_consulta.asp" -->
<%
response.buffer=false

if request("excel")=1 then
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
<!--#include file="../conn_consulta.asp" -->
<html>
<%
Server.ScriptTimeout=99999999

set db=server.createobject("ADODB.CONNECTION")

db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

aglutinado= left(formatnumber(request("selAglu")), len(formatnumber(request("selAglu")))-3)

set fonte=db.execute("SELECT * FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & aglutinado & " ORDER BY AGLU_SG_AGLUTINADO")

txtaglu=fonte("AGLU_SG_AGLUTINADO")

set rs=db.execute("SELECT * FROM SUB_MODULO WHERE SUMO_NR_CD_SEQUENCIA <> 33 AND SUMO_NR_CD_SEQUENCIA <> 34 AND SUMO_NR_CD_SEQUENCIA <> 36 ORDER BY SUMO_TX_DESC_SUB_MODULO")
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
    <td width="6%"><div align="right"><a href="gera_modulo_n_apoiado.asp?excel=1&selAglu=<%=request("selAglu")%>" target="_blank"><img src="../excel.jpg" width="22" height="20" border="0"></a></div></td>
    <td width="53%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Exportar 
      para o Excel</font></strong></td>
  </tr>
</table>
<p>
  <%end if%>
</p>
<table width="77%" border="0">
  <tr> 
    <td width="86%"><font color="#000066" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Relat&oacute;rio 
      de Assuntos n&atilde;o Apoiados por &Oacute;rg&atilde;o </strong></font></td>
    <td width="14%">&nbsp;</td>
  </tr>
</table>
<table width="77%" border="0">
  <tr> 
    <td width="72%" height="63" rowspan="2"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Oacute;rg&atilde;o 
      Aglutinador Selecionado : <%=txtaglu%></strong></font></td>
    <td width="1%" valign="baseline"> <div align="center"></div></td>
    <td width="27%" rowspan="2" valign="middle">&nbsp;</td>
  </tr>
  <tr> 
    <td valign="top"> <div align="center"></div></td>
  </tr>
</table>
<table width="83%" border="0">
  <tr valign="middle" bgcolor="#000066"> 
    <td height="36" colspan="3"> <div align="center"> 
        <p><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;os</font></strong> 
        </p>
      </div></td>
    <td><div align="center"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Assuntos</font><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        Apoiados</font></strong></div></td>
    <td height="36"><div align="center"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Assuntos</font><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        N&atilde;o Apoiados</font></strong></div></td>
  </tr>
  <%
  do until fonte.eof=true
  SG_AGLU=fonte("AGLU_SG_AGLUTINADO")
  COR1="#BFBFBF"
  %>
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
	if org1.eof=true then
	%>
  <tr> 
    <td colspan="3" bgcolor="<%=COR1%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><%=SG_AGLU%></b></font></td>
    <td bgcolor="#DBDBDB"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> </td>
    <td bgcolor="#DBDBDB"><font color="#000033" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUMO_TX_DESC_SUB_MODULO")%></font></td>
  </tr>
  <%
  else
  %>
  <tr> 
    <td colspan="3" bgcolor="<%=COR1%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><%=SG_AGLU%></b></font></td>
    <td bgcolor="#DBDBDB"><font color="#000033" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUMO_TX_DESC_SUB_MODULO")%></font></td>
	<td bgcolor="#DBDBDB"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; </font></td>
  </tr>
  <%
	end if
	rs.movenext
	SG_AGLU=" "
	COR1="WHITE"
	loop
	%></tr>
  <%
  set maior=db.execute("SELECT * FROM ORGAO_MAIOR WHERE AGLU_CD_AGLUTINADO=" & aglutinado & " AND ORLO_CD_STATUS='A' ORDER BY ORLO_SG_ORG_LOT")
  
  if trim(maior("ORLO_SG_ORG_LOT"))=trim(fonte("AGLU_SG_AGLUTINADO")) then
  	tem_menor=1
  else
  	tem_menor=0 		
  end if
  
  do until maior.eof=true
  
	aglutinador=aglutinado
    lotacao = maior("ORLO_CD_ORG_LOT")
    sequencia = maior("ORLO_NR_ORDEM")
    
    prefixo_orgao = Right("000" & aglutinador, 2) & Right("0000" & lotacao, 3) & (Right("000" & sequencia, 2))
    prefixo_orgao = Trim(prefixo_orgao)
		
	if len(prefixo_orgao) < 7 then
    	prefixo_orgao = Right("000" & aglutinador, 2) & Right("00000" & lotacao, 3) & (Right("000" & sequencia, 3))
		prefixo_orgao = Trim(prefixo_orgao)
	end if

    %>
    <%
	SG_MAIOR=maior("ORLO_SG_ORG_LOT")
	
	COR2="#E2E2E2"
	
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
	
	if tem_menor=1 then
		ssql=ssql+"                      (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs("SUMO_NR_CD_SEQUENCIA")& ") AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = '" & prefixo_orgao & "00000000')"
	else
		ssql=ssql+"                      (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs("SUMO_NR_CD_SEQUENCIA")& ") AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '" & trim(prefixo_orgao) & "%')"	
	end if
	
	'response.write ssql

	set org2=db.execute(ssql)
	
	MARCA="X"

	if org2.eof=true then
	%>
  <tr> 
    <td width="7%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td colspan="2" bgcolor="<%=COR2%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=SG_MAIOR%></b></font></td>
    <td bgcolor="#DBDBDB"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; </font></td>
    <td bgcolor="#DBDBDB"><font color="#000033" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUMO_TX_DESC_SUB_MODULO")%></font></td>
    <%
	else
	%>
  <tr> 
    <td width="7%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    <td colspan="2" bgcolor="<%=COR2%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=SG_MAIOR%></b></font></td>
    <td bgcolor="#DBDBDB"><font color="#000033" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUMO_TX_DESC_SUB_MODULO")%></font></td>
    <td bgcolor="#DBDBDB"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; </font></td>
	<%
	END IF
	rs.movenext
	SG_MAIOR=""
	COR2="WHITE"
	loop
	%>
  </tr>
  <%
	if tem_menor=1 then
	set menor = db.execute("SELECT * FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR LIKE '" & prefixo_orgao & "%' ORDER BY ORME_SG_ORG_MENOR")
	'set menor = db.execute("SELECT * FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR = '" & prefixo_orgao & "00000000' ORDER BY ORME_SG_ORG_MENOR")
	
	do until menor.eof=true
	
	if trim(menor("ORME_CD_ORG_MENOR"))<>trim(prefixo_orgao & "00000000") then

	SG_MENOR=menor("ORME_SG_ORG_MENOR")
	COR3="#EAEAEA"
	
	org_menor=menor("ORME_CD_ORG_MENOR")
	
	if right(org_menor,5)="00000" and right(left(org_menor,10),3)<>"000" then
	
	org_menor = left(org_menor,10)
		
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
	ssql=ssql+"                      (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = " & rs("SUMO_NR_CD_SEQUENCIA")& ") AND (dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR LIKE '" & org_menor & "%')"
	
	set org3=db.execute(ssql)
	
	if org3.eof=true then
	%>
  <tr> 
    <td>&nbsp;</td>
    <td width="5%">&nbsp;</td>
    <td width="23%" bgcolor="<%=COR3%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=SG_MENOR%></font></td>
    <td width="35%" bgcolor="#DBDBDB"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; </font></td>
    <td width="30%" bgcolor="#DBDBDB"><font color="#000033" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUMO_TX_DESC_SUB_MODULO")%></font></td>
    <%
	else
	%>
  <tr> 
    <td>&nbsp;</td>
    <td width="5%">&nbsp;</td>
    <td width="23%" bgcolor="<%=COR3%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=SG_MENOR%></font></td>
    <td width="35%" bgcolor="#DBDBDB"><font color="#000033" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUMO_TX_DESC_SUB_MODULO")%></font></td>
    <td width="30%" bgcolor="#DBDBDB"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; </font></td>
	<%
	END IF
	rs.movenext
	SG_MENOR=""
	COR3="WHITE"
	loop
	end if
	%>
  </tr>
  <%
	end if
  	menor.movenext
	loop
	end if
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
