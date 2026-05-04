<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
opti=request("op")
%>
<html>
<head>
<title>Visualização</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>

<body link="#000000" vlink="#000000" alink="#000000">
<table width="80%" border="0">
  <tr> 
    <td width="55%"><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>CONSULTA 
        POR NOME</strong></font></div></td>
    <td width="45%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Localizar 
      Texto : <strong>CTRL + F</strong></font> </td>
  </tr>
</table>
<p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#000066"> 
  </font></strong></font></p>
<table width="80%" border="0">
  <tr> 
    <td width="35%" height="45"> 
      <div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
    <td width="65%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#000066">Selecione 
      o Modo de Visualiza&ccedil;&atilde;o</font></strong></font></td>
  </tr>
</table>
<table width="75%" border="0">
  <tr> 
    <td width="53%" height="31">&nbsp;</td>
    <td width="47%" valign="middle"><font color="#000066"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="por_nome.asp?op=<%=opti%>">Visualizar 
      no Monitor</a></font></font></td>
  </tr>
  <tr> 
    <td height="28">&nbsp;</td>
    <td valign="middle"><font color="#000066"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="por_nome.asp?op=<%=opti%>&excel=1" target="_blank">Exportar 
      para o Excel</a></font></font></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
