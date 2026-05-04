 

<%
Response.Buffer = TRUE
Response.ContentType = "application/vnd.ms-excel"

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

hora1=request("hora1")
hora2=request("hora2")

if hora1="" then
	hora1="00:00"
end if

if hora2="" then
	hora2="00:00"
end if

data1=request("data1")
data2=request("data2")

if data1="" then
	data1="01/01/1900"
end if

if data2="" then
	data2="12/12/2222"
end if

DATA_INICIO=YEAR(DATA1) & "-" & day(DATA1) & "-" & month(DATA1)
DATA_TERM=YEAR(DATA2) & "-" & day(DATA2) & "-" & month(DATA2)

compl=" WHERE ONDA_CD_ONDA<>4 AND (ATUA_DT_ATUALIZACAO > CONVERT(DATETIME,  '" & DATA_INICIO & " " & hora1 & ":00', 102) AND  ATUA_DT_ATUALIZACAO < CONVERT(DATETIME, '" & DATA_TERM & " " & hora2 & ":00', 102))"

if ordem="" then
	ordem="ATUA_DT_ATUALIZACAO, CENA_CD_CENARIO"
end if

ordem=" ORDER BY " & ordem

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO" & compl & ordem

'response.write ssql

SSQL1=SSQL

set rs=db.execute(ssql)

IF RS.EOF=TRUE THEN
	TEM=0
ELSE
	TEM=1
END IF
%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<form name="frm1" method="POST" action="">
  <p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="3">Relatório
  Cenário x Status por Período</font> </p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  <table border="0" width="94%" cellspacing="0" cellpadding="0">
    <%if tem=1 then%>
    <tr>
      <td width="23%" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Cenário</b></font></td>
      <td width="14%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Status</b></font></td>
      <td width="18%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Tipo</b></font></td>
      <td width="17%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Configuração</b></font></td>
      <td width="26%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Desenvolvimento</b></font></td>
      <td width="32%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Data</b></font></td>
      <td width="26%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Hora</b></font></td>
    </tr>
    <%end if%>
    <%do until rs.eof=true
    IF COR="#E4E4E4" THEN
    	COR="WHITE"
    ELSE
    	COR="#E4E4E4"
    END IF
    %>
    <tr>
      <td width="23%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=rs("CENA_CD_CENARIO")%>-<%=RS("CENA_TX_TITULO_CENARIO")%></font></td>
      <%
      If rs("CENA_TX_SITUACAO") = "DF" Then
			      ls_Situacao_Cenario = "DEFINIDO"
			   elseIf rs("CENA_TX_SITUACAO") = "EE" Then
			      ls_Situacao_Cenario = "EM ELABORAÇÃO"
		      elseIf rs("CENA_TX_SITUACAO") = "DS" Then
				      ls_Situacao_Cenario = "DESENHADO"
			   elseIf rs("CENA_TX_SITUACAO") = "PT" Then
				      ls_Situacao_Cenario = "PRONTO PARA TESTE"
				elseIf rs("CENA_TX_SITUACAO") = "TD" Then
				      ls_Situacao_Cenario = "TESTADO NO PED"
				elseIf rs("CENA_TX_SITUACAO") = "TQ" Then
				      ls_Situacao_Cenario = "TESTADO NO PEQ"
			   end if
      %>
      <td width="14%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=ls_Situacao_Cenario%></font></td>
      <%
      if rs("CENA_TX_SITU_DESENHO_TIPO")=1 THEN
      		SITUACAO="COM DESENVOLVIMENTO"
      ELSE
      IF rs("CENA_TX_SITU_DESENHO_TIPO")=2 THEN
      		SITUACAO="SEM DESENVOLVIMENTO"
      ELSE
      		SITUACAO="  "
      END IF
      END IF
      %>
      <td width="18%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO%></font></td>
      <%
      if rs("CENA_TX_SITU_DESENHO_CONF")=0 THEN
      		SITUACAO2="  "
      ELSE
      IF rs("CENA_TX_SITU_DESENHO_CONF")=1 THEN
      		SITUACAO2="CONFIGURAÇÃO CONCLUÍDA"
      ELSE
      		SITUACAO2="  "
      END IF
      END IF
      %>
      <td width="17%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO2%></font></td>
      <%
      if rs("CENA_TX_SITU_DESENHO_DESE")=0 THEN
      		SITUACAO3="  "
      ELSE
      IF rs("CENA_TX_SITU_DESENHO_DESE")=1 THEN
      		SITUACAO3="DESENVOLVIMENTO CONCLUÍDO"
      ELSE
      		SITUACAO3="  "
      END IF
      END IF
      
      dia=day(RS("ATUA_DT_ATUALIZACAO"))
      mes=month(RS("ATUA_DT_ATUALIZACAO"))
      ano=year(RS("ATUA_DT_ATUALIZACAO"))
      
      atual1=dia & "/" & mes & "/" & ano
      atual2=formatdatetime(RS("ATUA_DT_ATUALIZACAO"), 4)

      %>
      <td width="26%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO3%></font></td>
      <td width="32%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="2"><%=atual1%></font></td>
      <td width="26%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="2"><%=atual2%></font></td>
    </tr>
  
    <%
      rs.movenext
      loop
      %>
</table>
<%if tem=0 then%>
  <font color="#800000" face="Verdana" size="2"><b>Nenhum Registro Encontrado</b></font>
 <%end if%>
  </form>
<p></p>
</body>
</html>
