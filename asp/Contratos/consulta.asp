<%
Server.ScriptTimeOut = 99999999

set db = Server.CreateObject("AdoDB.Connection")
db.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("db.mdb")
db.CursorLocation = 3

set db2=server.createobject("ADODB.CONNECTION")
db2.Open "Provider=SQLOLEDB.1;server=S6000db21;pwd=cogest00;uid=cogest;database=cogest"
db2.cursorlocation=3

set rs = db.execute("SELECT DISTINCT CBI, DESCRICAO FROM DESCCBI ORDER BY CBI")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Relatório de Situação de Contratos</title>
</head>

<body link="#000080" vlink="#000080" alink="#000080">

<p><b><font face="Verdana">Selecione o CBI Desejado</font></b></p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="70%" id="AutoNumber1" height="30">
           <tr>
                      <td width="20%" height="30"><p align="right"><img border="0" src="Imagens/seta.jpg"></td>
                      <td width="3%" height="30">&nbsp;</td>
                      <td width="73%" height="30"><font face="Verdana" size="2"><a href="gera_consulta.asp?cbi=X">TODOS OS CBI´S</a></font></td>
           </tr>
		   <%
           do until rs.eof=true

           prefixo = ""
           
           set v1 = db.execute("SELECT * FROM GERENTE WHERE CONTRATO LIKE '" & rs("cbi") & "%' ORDER BY CHAVE")
		   set v2 = db.execute("SELECT * FROM FISCAL WHERE CONTRATO LIKE '" & rs("cbi") & "%' ORDER BY CHAVE")

		   if v1.eof=false or v2.eof=false then

           		set rs2 = db.execute("SELECT ORG FROM COMPRAS WHERE CBI='" & rs("cbi") & "'")
           
           		if rs2.eof=false then
    	       		prefixo = " - " & rs.fields(1).value & " - ORG. COMPRAS : " & rs2.fields(0).value
	       		else
	           		prefixo = " - " & rs.fields(1).value & " - ORG. COMPRAS : - "
           		end if           
           %>           
           <tr>
                      <td width="20%" height="30"><p align="right"><img border="0" src="Imagens/seta.jpg"></td>
                      <td width="3%" height="30"><p align="center">&nbsp;</td>
                      <td width="73%" height="30"><font face="Verdana"><a href="gera_consulta.asp?cbi=<%=rs("cbi")%>"><%=rs("cbi")%><%=prefixo%></a></font></td>
           </tr>
           <%
           		end if
           		rs.movenext
           loop
           %>
</table>

</body>

</html>