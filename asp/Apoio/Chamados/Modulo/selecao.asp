<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../conecta.asp" -->
<%
caso = request("op")

select case caso
case 1
	endereco = "total_geral_dia.asp?modulo="
case 2
	endereco = "total_status.asp?modulo="
case 3
	endereco = "total_dia_status.asp?modulo="
case 4
	endereco = "atendimento_diario.asp?modulo="
case 5
	endereco = "perfil_atendimento.asp?modulo="
end select

set rs1 = db.execute("SELECT DISTINCT EQUIPE FROM " & Session("Tabela") & " WHERE EQUIPE IN ('', 'BASIS-PETROBRAS', 'GESTÃO DE EMPREENDIMENTOS', 'MANUTENÇÃO/INSPEÇÃO', 'MES-TREINAMENTO', 'ORÇAMENTO','SERVIÇOS', 'TREINAMENTO') ORDER BY EQUIPE")
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

<p><b><font face="Verdana" size="2">Selecione o Módulo desejado</font></b></p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="64%" id="AutoNumber1" height="26">
           
<%
do until rs1.eof=true
if trim(rs1("EQUIPE"))="" then                      
	equipe = "ATENDENTE TI"
else
	equipe = RS1("EQUIPE")
end if

%>
<tr>           
           <td width="11%" height="26"><p align="center"><img border="0" src="../../../../imagens/b011.gif"></td>
           <td width="89%" height="26"><font face="Verdana" size="1"><a href="<%=endereco%><%=EQUIPE%>"><%=EQUIPE%></a></font></td>
</tr>
<%
rs1.movenext
loop
%>
</table>
</body>

</html>