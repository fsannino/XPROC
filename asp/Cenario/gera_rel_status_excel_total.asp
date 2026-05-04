 

<%
Response.Buffer = TRUE
Response.ContentType = "application/vnd.ms-excel"

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql="select * from " & Session("PREFIXO") & "mega_processo order by MEPR_TX_dESC_MEGA_PROCESSO"

set rs_fonte=db.execute(ssql)

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE ONDA_CD_ONDA<>4 ORDER BY MEPR_CD_MEGA_PROCESSO, ONDA_CD_ONDA , CENA_CD_CENARIO"

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
  Cenário x Status</font> </p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  <table border="0" width="100%" cellspacing="0" cellpadding="0">
    <%if tem=1 then%>
    <tr> 
      <td width="8%" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Onda</b></font></td>
      <td width="8%" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Mega-Processo</b></font></td>
      <td width="5%" bgcolor="#330099" align="center"> <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#FFFFFF"><b>Código</b></font></p>
        <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#FFFFFF"><b>Cenário</b></font></p></td>
      <td width="5%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Nome 
        Cenário</b></font></td>
      <td width="5%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Status</b></font></td>
      <td width="13%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Tem 
        Desenvolvimento?</b></font></td>
      <td width="9%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Configuração</b></font></td>
      <td width="12%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Desenvolvimento</b></font></td>
      <td width="9%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Assunto</b></font></td>
      <td width="15%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Empresa</b></font></td>
      <td width="11%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Escopo</b></font></td>
    </tr>
    <%end if%>
    <%
    do until rs_fonte.eof=true
    
    ssql=""
    ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE MEPR_CD_MEGA_PROCESSO=" & rs_fonte("MEPR_CD_MEGA_PROCESSO") & " ORDER BY MEPR_CD_MEGA_PROCESSO, ONDA_CD_ONDA , CENA_CD_CENARIO"
	
	SET RS=DB.EXECUTE(SSQL)    
    
    do until rs.eof=true
    IF COR="#E4E4E4" THEN
    	COR="WHITE"
    ELSE
    	COR="#E4E4E4"
    END IF
    %>
    <tr> 
      <%SET RS_ONDA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA=" & rs("onda_cd_onda"))%>
      <td width="8%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=RS_ONDA("ONDA_TX_DESC_ONDA")%></font></td>
      <td width="8%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=RS_FONTE("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
      <td width="5%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=rs("CENA_CD_CENARIO")%></font></td>
      <td width="5%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=RS("CENA_TX_TITULO_CENARIO")%></font></td>
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
      <td width="5%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=ls_Situacao_Cenario%></font></td>
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
      <td width="13%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO%></font></td>
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
      <td width="9%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO2%></font></td>
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
	  	if not Isnull(rs("SUMO_NR_CD_SEQUENCIA")) then
      str_SQL = ""
      str_SQL = str_SQL & " SELECT SUMO_TX_DESC_SUB_MODULO, "
      str_SQL = str_SQL & "     SUMO_NR_CD_SEQUENCIA"
      str_SQL = str_SQL & " FROM SUB_MODULO"
      str_SQL = str_SQL & " WHERE SUMO_NR_CD_SEQUENCIA = " & rs("SUMO_NR_CD_SEQUENCIA")
      set rs_Modulo = db.Execute(str_SQL)
      if not rs_Modulo.EOF then
         str_DsAssunto = rs_Modulo("SUMO_TX_DESC_SUB_MODULO")
      else
         str_DsAssunto = " não enconttado o assunto "
      end if
      rs_Modulo.close
	else
	  str_DsAssunto = ""
	end if  

    If rs("CENA_TX_SITUACAO_VALIDACAO") = "0" Then
       ls_Situacao_Escopo = "Fora Escopo"
    elseIf rs("CENA_TX_SITUACAO_VALIDACAO") = "1" then
	   ls_Situacao_Escopo = "No Escopo"
    end if    
      %>
      <td width="12%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO3%></font></td>
      <td width="9%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=str_DsAssunto%></font></td>
      <td width="15%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=RS("CENA_TX_EMPRESA_RELAC")%></font></td>
      <td width="11%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=ls_Situacao_Escopo%></font></td>
    </tr>
    <%
      rs.movenext
      loop
      rs_fonte.movenext
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
