<!--#include file="../conecta.asp" -->
<html>
<%
categ = request("modulo")

if categ = "ATENDENTE TI" then
	categ=""
end if

set rs1 = db.execute("SELECT DISTINCT EQUIPE FROM " & Session("Tabela") & " WHERE EQUIPE='" & categ & "' ORDER BY EQUIPE")

set rs = db.execute("SELECT DISTINCT SITUACAO FROM " & Session("tabela") & " ORDER BY SITUACAO")
%>
<head>
</head>

<script>
function carrega_correio()
{
alert('Funcionalidade ainda não implementada...');
//window.location='correio.asp'
}
</script>

<body>

<p><b><font face="Verdana" size="2">Total Geral de Registros por Situação - Por Módulo</font></b></p>

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
           </tr>
           <%
           do until rs.eof=true
           rs1.movefirst
           %>
           <tr>
                   <td width="42%" height="29" align="center"><font face="Verdana" size="1"><%=rs("situacao")%></font></td>
			<%
			do until rs1.eof=true
			
            data_01 = cdate(session("data_inicio"))
       	    data_inicio = year(data_01) & "-" & right("000" & month(data_01),2) & "-" & right("000" & day(data_01),2)

			ssql=""
			ssql="SELECT * FROM " & Session("tabela") & " WHERE SITUACAO='" & rs("situacao") & "' AND EQUIPE ='" & rs1("EQUIPE") & "'"
			ssql=ssql+" AND (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102))"
			ssql=ssql+ Session("compl")
			
			set rs2 = db.execute(ssql)
           
            itens = rs2.recordcount
 			%>                   
		       <td width="44%" height="29" align="center"><font face="Verdana" size="1"><%=itens%></font></td>
		    <%
		    rs1.movenext
		    loop
		    %>
           </tr>
           <%
           rs.movenext
           loop
           %>
</table>
<p><input type="button" value="Enviar Correio" name="B3" onClick="carrega_correio()"></p>
</body>

</html>