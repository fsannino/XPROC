<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%>
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

mega = request("mega")
processo = request("processo")

set rs=conn_db.execute("SELECT * FROM SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & MEGA & " AND PROC_CD_PROCESSO=" & PROCESSO & " ORDER BY SUPR_TX_DESC_SUB_PROCESSO")

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Histórico de Validação</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#990000" vlink="#990000" alink="#990000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="http://its_server3/valida_status1.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1">
  <table width="64%" border="0">
    <tr>
      <td width="67%"><font color="#330099" face="Verdana" size="3">Visualização Sub-Processos</font></td>
      <td width="33%"><div align="right"><strong><font color="#990000" size="3"><a href="javascript:window.close()">Fechar 
          Janela</a></font></strong></div></td>
    </tr>
  </table>
  <br>
  <table width="64%" border="0">
    <tr bgcolor="#000066"> 
      <td width="70%"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Sub-Processo</font></strong></td>
      <td width="30%"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Impacto</font></strong></td>
    </tr>
  <%
  tem=0
  do until rs.eof=true
  %>
  <tr> 
      <td><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUPR_TX_DESC_SUB_PROCESSO")%></font></td>
      <%
	  SELECT CASE rs("SUPR_TX_IMPACTO")
	  CASE 1
		IMPACTO = "Alto"	  	  
	  CASE 2
		IMPACTO = "Médio"	  	  
	  CASE 3
		IMPACTO = "Baixo"	  	  
	  CASE ELSE
		IMPACTO = "Não Definido"	  
	  END SELECT
	  %>
	  <td><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=IMPACTO%></font></td>
  </tr>
  <%
  tem=tem+1
  rs.movenext
  loop
  %>
  </table>
  <%if tem=0 then%>
  <p><font color="#990000"><strong>Nenhum Registro Encontrado</strong></font></p>
  <%end if%>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%>
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

mega = request("mega")
processo = request("processo")

set rs=conn_db.execute("SELECT * FROM SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & MEGA & " AND PROC_CD_PROCESSO=" & PROCESSO & " ORDER BY SUPR_TX_DESC_SUB_PROCESSO")

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Histórico de Validação</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#990000" vlink="#990000" alink="#990000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="http://its_server3/valida_status1.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1">
  <table width="64%" border="0">
    <tr>
      <td width="67%"><font color="#330099" face="Verdana" size="3">Visualização Sub-Processos</font></td>
      <td width="33%"><div align="right"><strong><font color="#990000" size="3"><a href="javascript:window.close()">Fechar 
          Janela</a></font></strong></div></td>
    </tr>
  </table>
  <br>
  <table width="64%" border="0">
    <tr bgcolor="#000066"> 
      <td width="70%"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Sub-Processo</font></strong></td>
      <td width="30%"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Impacto</font></strong></td>
    </tr>
  <%
  tem=0
  do until rs.eof=true
  %>
  <tr> 
      <td><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUPR_TX_DESC_SUB_PROCESSO")%></font></td>
      <%
	  SELECT CASE rs("SUPR_TX_IMPACTO")
	  CASE 1
		IMPACTO = "Alto"	  	  
	  CASE 2
		IMPACTO = "Médio"	  	  
	  CASE 3
		IMPACTO = "Baixo"	  	  
	  CASE ELSE
		IMPACTO = "Não Definido"	  
	  END SELECT
	  %>
	  <td><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=IMPACTO%></font></td>
  </tr>
  <%
  tem=tem+1
  rs.movenext
  loop
  %>
  </table>
  <%if tem=0 then%>
  <p><font color="#990000"><strong>Nenhum Registro Encontrado</strong></font></p>
  <%end if%>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
