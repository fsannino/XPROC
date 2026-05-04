<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../conecta.asp" -->
<%
caso = request("op")

select case caso
case 1
	endereco = "total_geral_dia.asp?categoria="
case 2
	endereco = "total_status.asp?categoria="
case 3
	endereco = "total_dia_status.asp?categoria="
case 4
	endereco = "atendimento_diario.asp?categoria="
case 5
	endereco = "perfil_atendimento.asp?categoria="
end select

ssql="SELECT DISTINCT CATEGORIA FROM " & Session("Tabela") & " ORDER BY CATEGORIA"

set rs1 = db.execute(ssql)
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Seleção</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

</head>

<body link="#00509F" vlink="#00509F" alink="#00509F">

<p><b><font face="Verdana" size="2">Selecione o Grupo de Solucionadores desejado</font></b></p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="64%" id="AutoNumber1" height="26">
           
<%
do until rs1.eof=true
%>
<tr>           
           <td width="11%" height="26"><p align="center"><img border="0" src="../../../../imagens/b011.gif"></td>
           <td width="89%" height="26"><font face="Verdana" size="1"><a href="<%=endereco%><%=rs1("CATEGORIA")%>"><%=rs1("CATEGORIA")%></a></font></td>
</tr>
<%
rs1.movenext
loop
%>
</table>
</body>

</html>