<%
Server.ScriptTimeOut = 99999999

set db = Server.CreateObject("AdoDB.Connection")
db.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("db.mdb")
db.CursorLocation = 3
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Relação de Excessões</title>
</head>

<body link="#000080" vlink="#000080" alink="#000080">
<%
set rs = db.execute("SELECT DISTINCT CONTRATO FROM CONTRATO ORDER BY CONTRATO")

i = 0
reg = rs.recordcount

do until i = reg
	
	set temp1 = db.execute("SELECT * FROM FISCAL WHERE CONTRATO='" & rs("CONTRATO") & "'")
	set temp2 = db.execute("SELECT * FROM GERENTE WHERE CONTRATO='" & rs("CONTRATO") & "'")
	
	if temp1.eof=false and temp2.eof=true then
		gerente = gerente & rs("CONTRATO") & ", "	
	end if
	
	if temp1.eof=true and temp2.eof=false then
		fiscal = fiscal & rs("CONTRATO") & ", "	
	end if	
	
	i = i + 1
	
	rs.movenext
loop

if len(fiscal)>2 then
	fiscal = left(fiscal, len(fiscal)-2)
else
	fiscal = "Nenhum Registro Encontrado"
end if

if len(gerente)>2 then
	gerente = left(gerente, len(gerente)-2)
else
	gerente = "Nenhum Registro Encontrado"
end if


%>
<p><b><font face="Verdana" color="#000080">Contratos com Gerente, mas sem Fiscal</font></b></p>
<p><font face="Verdana" size="2"><%=fiscal%></font></p>
<p>&nbsp;</p>
<p><b><font face="Verdana" color="#000080">Contratos com Fiscal, mas sem Gerente</font></b></p>
<p><font face="Verdana" size="2"><%=gerente%></font></p>

</body>

</html>