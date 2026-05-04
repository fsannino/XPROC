<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../conecta.asp" -->
<%
categ = request("modulo")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Nova pagina 1</title>
</head>

<body>

<p><b><font face="Verdana" size="2">Perfil de Atendimento - Por Módulo</font></b></p>

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

<p><font face="Verdana" size="2">Módulo Selecionado :<b> <%=categ%></b></font></p>
<%
if categ="ATENDENTE TI" then
	categ=""
end if
%>
<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#758A8A" width="44%" id="AutoNumber1" height="53">
           <tr>
				<td width="41%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Data</font></b></td>
                <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Total de Registros no dia</font></b></td>
                <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Fechados em até 1 hora</font></b></td>
                <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Fechados em até 1 dia</font></b></td>
                <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Fechados em até 3 dias</font></b></td>
                <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Fechados </font></b></td>
                <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Não Fechados</font></b></td>
           </tr>
           <%
           do until i = Session("Periodo")
           
	           data_01 = cdate(session("data_inicio")) + i
		       data_02 = cdate(session("data_inicio")) + (i + 1)
           
        	   data_inicio = year(data_01) & "-" & right("000" & month(data_01),2) & "-" & right("000" & day(data_01),2)
        	   data_fim = year(data_02) & "-" & right("000" & month(data_02),2) & "-" & right("000" & day(data_02),2)

	           ssql=""
	           ssql="SELECT REGISTRO, ABERTURA "
	           ssql=ssql+"FROM         dbo." & session("tabela")
	           ssql=ssql+" WHERE      (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102)) AND (ABERTURA < CONVERT(DATETIME, '" & data_fim & " 00:00:00', 102))"
	           ssql=ssql+" AND EQUIPE='" & categ & "'"
	           ssql=ssql+ Session("compl")
          
    	       set rs = db.execute(ssql)
           
	           itens = rs.recordcount

    	       'ate 1 hora
    	       
			   ssql=""
	           ssql="SELECT REGISTRO, ABERTURA "
	           ssql=ssql+"FROM         dbo." & session("tabela")
	           ssql=ssql+" WHERE      (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102)) AND (ABERTURA < CONVERT(DATETIME, '" & data_fim & " 00:00:00', 102))"
	           ssql=ssql+"AND DATEDIFF (MINUTE , ABERTURA , SOLUCAO) < 60"
	           ssql=ssql+" AND EQUIPE='" & categ & "'"
	           ssql=ssql+ Session("compl")	           

    	       set rs = db.execute(ssql)
    	       
    	       ate01h = rs.RecordCount

    	       'ate 1 dia
    	       
			   ssql=""
	           ssql="SELECT REGISTRO, ABERTURA "
	           ssql=ssql+"FROM         dbo." & session("tabela")
	           ssql=ssql+" WHERE      (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102)) AND (ABERTURA < CONVERT(DATETIME, '" & data_fim & " 00:00:00', 102))"
	           ssql=ssql+"AND DATEDIFF (HOUR , ABERTURA , SOLUCAO) < 24"
	           ssql=ssql+" AND EQUIPE='" & categ & "'"
	           ssql=ssql+ Session("compl")	           
	           

    	       set rs = db.execute(ssql)
    	       
    	       ate01d = rs.RecordCount

    	       'ate 3 dias
    	       
			   ssql=""
	           ssql="SELECT REGISTRO, ABERTURA "
	           ssql=ssql+"FROM         dbo." & session("tabela")
	           ssql=ssql+" WHERE      (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102)) AND (ABERTURA < CONVERT(DATETIME, '" & data_fim & " 00:00:00', 102))"
	           ssql=ssql+"AND DATEDIFF (HOUR , ABERTURA , SOLUCAO) < 72"
	           ssql=ssql+" AND EQUIPE='" & categ & "'"
	           ssql=ssql+ Session("compl")	           

    	       set rs = db.execute(ssql)
    	       
    	       ate03d = rs.RecordCount

    	       'ate 3 dias
    	       
			   ssql=""
	           ssql="SELECT REGISTRO, ABERTURA "
	           ssql=ssql+"FROM         dbo." & session("tabela")
	           ssql=ssql+" WHERE      (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102)) AND (ABERTURA < CONVERT(DATETIME, '" & data_fim & " 00:00:00', 102))"
	           ssql=ssql+"AND DATEDIFF (HOUR , ABERTURA , SOLUCAO) > 72"
	           ssql=ssql+" AND EQUIPE='" & categ & "'"
	           ssql=ssql+ Session("compl")	           

    	       set rs = db.execute(ssql)
    	       
    	       mais03d = rs.RecordCount

	       if session("Modo") = "Q" then
		   %>
           <tr>
                      <td width="41%" height="29" align="center"><font face="Verdana" size="1"><%=data_01%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=itens%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=ate01h%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=ate01d%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=ate03d%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=ate03d + mais03d%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=itens - (ate03d + mais03d)%></font></td>
           </tr>
           <%
		   else
		   %>
           <tr>
                      <td width="41%" height="29" align="center"><font face="Verdana" size="1"><%=data_01%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=itens%></font></td>
                      <%
	         			if itens = 0 then
					  		itens=1
						end if
					  %>
					  <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=formatpercent((ate01h/itens))%> </font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=formatpercent((ate01d/itens))%> </font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=formatpercent((ate03d/itens))%> </font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=formatpercent(((ate03d + mais03d)/itens))%> </font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=formatpercent(((itens - (ate03d + mais03d))/itens))%></font></td>
           </tr>

		   <%
		   end if
           i = i + 1
	       loop
    	   %>
</table>
</body>

</html>