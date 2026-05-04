<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conecta.asp" -->
<%
server.scripttimeout=99999999
response.buffer=false

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

cod= request("Registro")

ssql=""
ssql="SELECT *"
ssql=ssql+" FROM BACKLOG WHERE BALO_CD_COD_BACKLOG=" & COD
ssql=ssql+" ORDER BY BALO_TX_TITULO"

set rs = db.execute(ssql)

%>
<html>
<head>
<title>#BACKLOG - Solicitações de Melhoria no SAP R/3#</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0033FF" alink="#0033FF" link="#0033FF">
<p>&nbsp;</p>
<p align="center"><font face="Verdana" color="#000080">Consulta de Solicita&ccedil;&otilde;es 
  de Melhoria na Solu&ccedil;&atilde;o Configurada no SAP R/3</font> </p>
<p align="center">&nbsp;</p>
<%
do until rs.eof=true
%>
<table width="86%" border="0" height="369">
  <tr> 
    <td width="20%" height="36">&nbsp;</td>
    <td width="24%" bgcolor="#000099" height="36"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">Titulo 
      da Solicita&ccedil;&atilde;o</font></font></b></td>
    <td width="56%" height="36"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("BALO_TX_TITULO")%></font></b></td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="24%" bgcolor="#000099"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">Mega-Processo</font></font></b></td>
	<%
	set temp = db.execute("SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))	
	%>
    <td width="56%"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></font></b></td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="24%" bgcolor="#000099"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">Assunto</font></font></b></td>
    <%
	if rs("SUMO_NR_CD_SEQUENCIA")<>0 THEN
		SET TEMP = DB.EXECUTE("SELECT * FROM SUB_MODULO WHERE SUMO_NR_CD_SEQUENCIA=" & rs("SUMO_NR_CD_SEQUENCIA"))
		T_ASSUNTO = TEMP("SUMO_TX_DESC_SUB_MODULO")
	ELSE
		T_ASSUNTO=""
	END IF
	%>
    <td width="56%"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=T_ASSUNTO%></font></b></td>
  </tr>
  <tr> 
    <td width="20%" height="28">&nbsp;</td>
    <td width="24%" bgcolor="#000099" height="28"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">&Oacute;rg&atilde;o</font></font></b></td>
	<%
	IF LEN(RS("ORME_CD_ORG_MENOR"))=2 THEN
		SET TEMP = DB.EXECUTE("SELECT DISTINCT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & RS("ORME_CD_ORG_MENOR"))
	ELSE
		SET TEMP = DB.EXECUTE("SELECT DISTINCT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & RS("ORME_CD_ORG_MENOR") & "'")	
	END IF
	%>
    <td width="56%" height="28"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=TEMP("ORGAO")%></font></b></td>
  </tr>
  <tr> 
    <td width="20%" height="89">&nbsp;</td>
    <td width="24%" height="89" bgcolor="#000099" valign="top"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">Descri&ccedil;&atilde;o 
      da Solicita&ccedil;&atilde;o</font></font></b></td>
    <td width="56%" height="89" valign="top"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("BALO_TX_DESCRICAO")%></font></b></td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="24%" bgcolor="#000099"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF"> 
      Solicitante</font></font></b></td>
    <td width="56%"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=rs("BALO_TX_CHAVE")%> - <%=rs("BALO_TX_SOLICITANTE")%> - <%=rs("BALO_TX_TELEFONE")%></font></b></td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="24%" bgcolor="#000099"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">Respons&aacute;vel 
      no Sinergia</font></font></b></td>
    <%
	  SET TEMP = DB.EXECUTE("SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO =" & RS("BALO_CD_RESPONSAVEL"))
	  T_RESP = TEMP("MEPR_TX_ABREVIA")
	  %>
    <td width="56%"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=T_RESP%></font></b></td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="24%" bgcolor="#000099"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">Prioridade</font></font></b></td>
    <%
	SELECT CASE RS("BALO_CD_PRIORIDADE")
	CASE 1
		T_PRIOR = "PRÉ GOLIVE SINERGIA"
	CASE 2
		T_PRIOR = "PÓS GOLIVE"		
	END SELECT
	%>
    <td width="56%"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=T_PRIOR%></font></b></td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="24%" bgcolor="#000099"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">Tipo</font></font></b></td>
    <%
	SELECT CASE RS("BALO_CD_TIPO")
	CASE 1
		T_TIPO = "CORRETIVA"
	CASE 2
		T_TIPO = "LEGAL / NORMATIVA"		
	CASE 3
		T_TIPO = "MELHORIA"		
	END SELECT
	%>
    <td width="56%"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=T_TIPO%></font></b></td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="24%" bgcolor="#000099"><b><font size="2" face="Arial, Helvetica, sans-serif"><font color="#FFFFFF">Existe 
      no Legado?</font></font></b></td>
    <%
	SELECT CASE RS("BALO_CD_LEGADO")
	CASE 1
		T_LEG = "SIM"
	CASE 2
		T_LEG = "NÃO"		
	END SELECT
	%>
    <td width="56%"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=T_LEG%></font></b></td>
  </tr>
</table>
<%
tem=tem+1
rs.movenext
loop
%>
<p align="center"><b><font color="#0033FF" face="Courier New, Courier, mono" size="2"><a href="javascript:history.go(-1)">Retornar 
  para a Tela Anterior</a></font></b></p>
<p align="center">&nbsp;</p>
</body>
</html>