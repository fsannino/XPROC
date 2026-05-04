<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conecta.asp" -->
<%
registro = request("Registro")

server.scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

ssql = "DELETE FROM BACKLOG WHERE BALO_CD_COD_BACKLOG=" & registro

db.execute(ssql)

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

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#000099" alink="#000099" link="#000099">
<p>&nbsp;</p>
<p align="center"><font face="Verdana" color="#000080">Solicita&ccedil;&otilde;es 
  de Melhoria na Solu&ccedil;&atilde;o Configurada no SAP R/3</font> </p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Verdana" color="#000080"><b><font size="2">O Registro 
  foi Exclu&iacute;do com Sucesso!</font></b></font></p>
<table width="75%" border="0" align="center">
  <tr> 
    <td width="30%" height="58">&nbsp;</td>
    <td width="6%" height="58"> 
      <div align="center"><img src="seta_d.jpg" width="23" height="24"></div>
    </td>
    <td width="64%" height="58"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099"><a href="index.asp">Retornar 
      ao Menu Principal</a></font></td>
  </tr>
</table>
<p align="center">&nbsp;</p>
</body>
</html>