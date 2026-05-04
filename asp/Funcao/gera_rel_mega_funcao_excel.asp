 
<%
Response.Buffer = TRUE
Response.ContentType = "application/vnd.ms-excel"

server.scripttimeout=99999999
on error resume next

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")
func=request("selFuncao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE FUNE_CD_FUNCAO_NEGOCIO='"& func &"'")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0" link="#000000" vlink="#000000" alink="#000000">

<p style="margin-top: 0; margin-bottom: 0">

&nbsp;&nbsp;&nbsp;<font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">
Relatório Fun&ccedil;&atilde;o R/3 x Transa&ccedil;&atilde;o</font> 
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)%>
<p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="2"><b>&nbsp;&nbsp;&nbsp;
Mega-Processo
: </b><%=mega%> - <%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></font></p>
<%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & func & "'")%>
<p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="2"><b>&nbsp;&nbsp;&nbsp;
Fun&ccedil;&atilde;o R/3 : </b><%=FUNC%> - <%=TEMP("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%IF RS.EOF=FALSE THEN%>
<table border="0" cellspacing="1" cellpadding="2" width="870" bordercolor="#000000">
  <tr>
    <td width="159" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Mega-Processo</font></b></td>
    <td width="172" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Processo</font></b></td>
    <td width="185" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Sub-Processo</font></b></td>
    <td width="204" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Atividade</font></b></td>
    <td width="138" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Transação</font></b></td>
  </tr>
  <%DO UNTIL RS.EOF=TRUE%>
  <tr>
    <td width="159">
      <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO"))%>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
    <td width="172">
      <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO=" & RS("PROC_CD_PROCESSO"))%>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=TEMP("PROC_TX_DESC_PROCESSO")%></font></td>
    <td width="185">
    <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO=" & RS("PROC_CD_PROCESSO") & " AND SUPR_CD_SUB_PROCESSO=" & RS("SUPR_CD_SUB_PROCESSO"))%>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=TEMP("SUPR_TX_DESC_SUB_PROCESSO")%></font></td>
    <td width="204">
    <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & RS("ATCA_CD_ATIVIDADE_CARGA"))%>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=TEMP("ATCA_TX_DESC_ATIVIDADE")%></font></td>
    <td width="138">
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=RS("TRAN_CD_TRANSACAO")%></font></td>
  </tr>
  <%
  RS.MOVENEXT
  LOOP
  %>
</table>
<%ELSE%>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;<b><font face="Verdana" size="2" color="#800000">
Nenhum
Registro Encontrado</font></b></p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<%END IF%>

</body>

</html>