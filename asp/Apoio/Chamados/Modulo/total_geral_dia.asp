<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../conecta.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Nova pagina 1</title>
</head>

<body>
<%
categ = request("modulo")

if categ = "ATENDENTE TI" then
	categ=""
end if

set rs1 = db.execute("SELECT DISTINCT EQUIPE FROM " & Session("Tabela") & " WHERE EQUIPE='" & categ & "' ORDER BY EQUIPE")
%>

<p><b><font face="Verdana" size="2">Total Geral de Registros por dia - Por Módulo</font></b></p>

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
<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#758A8A" width="44%" id="AutoNumber1" height="53">
           <tr>
                      <td width="41%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Data</font></b></td>
                      <%
                      do until rs1.eof=true
                      if trim(rs1("EQUIPE"))="" then                      
                      	equipe="ATENDENTE TI"
                      else
                      	equipe=RS1("EQUIPE")
                      end if
                      %>
		              <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2"><%=EQUIPE%></font></b></td>
		              <%
		              rs1.movenext
		              loop
		              %>
		              <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">TOTAL DE REGISTROS</font></b></td>
           </tr>
           <%
           do until i = Session("Periodo")
           
           data_01 = cdate(session("data_inicio")) + i
           data_02 = cdate(session("data_inicio")) + (i + 1)
           
           data_inicio = year(data_01) & "-" & right("000" & month(data_01),2) & "-" & right("000" & day(data_01),2)
           data_fim = year(data_02) & "-" & right("000" & month(data_02),2) & "-" & right("000" & day(data_02),2)
          
           %>
           
           <font size="1" face="Verdana"> </font>
           <tr>
                      <td width="41%" height="29" align="center"><font face="Verdana" size="1"><%=data_01%></font></td>
           <%
           rs1.movefirst
		   total = 0
           
           do until rs1.eof=true
			           ssql=""
			           ssql="SELECT DISTINCT TOP 100 PERCENT REGISTRO, ABERTURA "
			           ssql=ssql+"FROM         dbo." & session("tabela")
			           ssql=ssql+" WHERE     (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102) AND ABERTURA < CONVERT(DATETIME, '" & data_fim & " 00:00:00', 102))"
			           ssql=ssql+" AND EQUIPE='" & RS1("EQUIPE") & "'"
			           ssql=ssql+ Session("compl")
          
			           set rs = db.execute(ssql)
           
			           itens = rs.recordcount
			           
			           total = total + itens
			           %>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=itens%></font></td>
			<%
			rs1.movenext
			loop
			%>
            <td width="45%" height="29" align="center" bgcolor="#E2E2E2"><font face="Verdana" size="1"><%=total%></font></td>
			</tr>
			<%
			i = i + 1
			loop
           %>
</table>
</body>

</html>