<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="conecta.asp" -->
<html>
<%
set rs = db.execute("SELECT DISTINCT SITUACAO FROM " & Session("tabela") & " ORDER BY SITUACAO")
%>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Nova pagina 1</title>
</head>

<body>

<p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Total Geral de Registros por Situação</font></b></p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="56%" id="AutoNumber1" height="26">
           <tr>
                      <td width="48%" height="1" align="center" bgcolor="#E5E5E5"><p align="left"><font size="1" color="#000080"><b><font face="Verdana">Data Base Inicial : </font></b><font face="Verdana"><%=Session("data_inicio")%></font></font></td>
                      <td width="52%" height="1" align="center" bgcolor="#E5E5E5"><p align="left"><font size="1" color="#000080"><b><font face="Verdana">Período : </font></b><font face="Verdana"><%=Session("periodo")%> dias</font></font></td>
           </tr>
           <tr>
                      <td width="46%" height="1" align="center" bgcolor="#E5E5E5"><p align="left"><b><font face="Verdana" size="1" color="#000080">Tipo</font></b><font face="Verdana" size="1" color="#000080"><b> : </b><%=Session("Erro")%></font></td>
                      <td width="54%" height="1" align="center" bgcolor="#E5E5E5"><p align="left"><font face="Verdana" size="1" color="#000080"><b>Órgão : </b><%=Session("Orgao")%></font></td>
           </tr>
</table>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#758A8A" width="45%" id="AutoNumber1" height="53">
           <tr>
                      <td width="42%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Situação</font></b></td>
                      <td width="44%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Total de Registros</font></b></td>
           </tr>
           <%
           do until rs.eof=true
           
           data_01 = cdate(session("data_inicio"))
           data_inicio = year(data_01) & "-" & right("000" & month(data_01),2) & "-" & right("000" & day(data_01),2)
       	   
       	   ssql=""
       	   ssql=ssql+"SELECT * FROM " & Session("tabela") & " WHERE SITUACAO='" & rs("situacao") & "' "
       	   ssql=ssql+" AND (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102))"
       	   ssql=ssql+ Session("compl")
       	   
           set rs1 = db.execute(ssql)
           
           itens = rs1.recordcount
           %>
           <tr>
                      <td width="42%" height="29" align="center"><font face="Verdana" size="1"><%=rs("situacao")%></font></td>
                      <td width="44%" height="29" align="center"><font face="Verdana" size="1"><%=itens%></font></td>
           </tr>
           <%
           rs.movenext
           loop
           %>
</table>
</body>

</html>