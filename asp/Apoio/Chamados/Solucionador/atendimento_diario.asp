<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../conecta.asp" -->
<%
categ = request("categoria")
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

<p><b><font face="Verdana" size="2">Atendimento Diário - Por Grupo de Solucionadores</font></b></p>

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

<p><font face="Verdana" size="2">Solucionador Selecionado :<b> <%=categ%></b></font></p>

<table border="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#758A8A" width="44%" id="AutoNumber1" height="53">
           <tr>
                      <td width="41%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Data</font></b></td>
                      <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Saldo no Início do dia</font></b></td>
                      <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Abertos no Dia</font></b></td>
                      <td width="45%" height="23" bgcolor="#758A8A" align="center"><b><font face="Verdana" size="1" color="#E2E2E2">Fechados no Dia</font></b></td>

           </tr>
           <%
           do until i = Session("Periodo")
           
	           data_01 = cdate(session("data_inicio")) + i
		       data_02 = cdate(session("data_inicio")) + (i + 1)
           
        	   data_inicio = year(data_01) & "-" & right("000" & month(data_01),2) & "-" & right("000" & day(data_01),2)
        	   data_fim = year(data_02) & "-" & right("000" & month(data_02),2) & "-" & right("000" & day(data_02),2)
           
if i = 0 then
	
	abertos=0
	fechados=0

	ssql=""
	ssql="SELECT REGISTRO, ABERTURA "
	ssql=ssql+"FROM         dbo." & session("tabela")
	ssql=ssql+" WHERE     (ABERTURA < CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102))"
	ssql=ssql+" AND CATEGORIA='" & categ & "'"
	ssql=ssql+ Session("compl")
          
	set rs = db.execute(ssql)
    	       
    	abertos = rs.recordcount

	ssql=""
	ssql="SELECT REGISTRO, ABERTURA "
	ssql=ssql+"FROM         dbo." & session("tabela")
	ssql=ssql+" WHERE     (SOLUCAO < CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102))"
	ssql=ssql+" AND SITUACAO='FECHADO' OR SITUACAO='CANCELADO'"
	ssql=ssql+" AND CATEGORIA='" & categ & "'"
	ssql=ssql+ Session("compl")
          
	set rs = db.execute(ssql)
		       
    	fechados = rs.recordcount
             	
	itens = abertos - fechados

else

	itens = (itens + abertos) - fechados

end if			
		       		       
		       if itens < 0 then
		       		itens = 0
		       end if	
		       
		       abertos=0
		       fechados=0

	           ssql=""
	           ssql="SELECT DISTINCT REGISTRO, ABERTURA "
	           ssql=ssql+"FROM         dbo." & session("tabela")
	           ssql=ssql+" WHERE     (ABERTURA > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102) AND ABERTURA < CONVERT(DATETIME, '" & data_fim & " 00:00:00', 102))"
	           ssql=ssql+" AND CATEGORIA='" & categ & "'"
	           ssql=ssql+ Session("compl")
          
    	       set rs = db.execute(ssql)
    	       
    	       abertos = rs.recordcount

	           ssql=""
	           ssql="SELECT DISTINCT REGISTRO, ABERTURA "
	           ssql=ssql+"FROM         dbo." & session("tabela")
	           ssql=ssql+" WHERE     (SOLUCAO > CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102) AND SOLUCAO < CONVERT(DATETIME, '" & data_fim & " 00:00:00', 102))"
	           ssql=ssql+" AND SITUACAO='FECHADO'"
	           ssql=ssql+" AND CATEGORIA='" & categ & "'"
	           ssql=ssql+ Session("compl")
          
    	       set rs = db.execute(ssql)
		       
       	       fechados = rs.recordcount
       	       
		       %>
           
           <font size="1" face="Verdana"> </font>
           <tr>
                      <td width="41%" height="29" align="center"><font face="Verdana" size="1"><%=data_01%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=itens%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=abertos%></font></td>
                      <td width="45%" height="29" align="center"><font face="Verdana" size="1"><%=fechados%></font></td>
           </tr>
        	   <%
	           i = i + 1
	       loop
    	   %>
</table>
</body>

</html>