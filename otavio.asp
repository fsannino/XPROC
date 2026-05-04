<% 

str_String_Comm= "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"

set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open str_String_Comm
conn_Cogest.CursorLocation=3

str_SQL = ""
str_SQL = str_SQL & " SELECT  "
str_SQL = str_SQL & " LOTE_NR_SEQ_LOTE"
str_SQL = str_SQL & " ,LOTE_TX_DESCRICAO"
str_SQL = str_SQL & " , LOTE_DT_ENVIO"
str_SQL = str_SQL & " , LOTE_NR_QTD_EXPORTACAO"
str_SQL = str_SQL & " , LOTE_TX_ORGAO_SELEC"
str_SQL = str_SQL & " , LOTE_TX_FUNCAO_SELEC"
str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
str_SQL = str_SQL & " FROM dbo.GOLI_LOTE"
str_SQL = str_SQL & " order by LOTE_NR_SEQ_LOTE desc"

set rds_Lote = conn_Cogest.Execute(str_SQL)

%>

<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="">
  <table width="100%"  border="0" cellspacing="0" cellpadding="1">
    <tr>
      <td width="21%">&nbsp;</td>
      <td width="46%">&nbsp;</td>
      <td width="33%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Rela&ccedil;&atilde;o de Lotes</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="87%"  border="0" cellspacing="5" cellpadding="1">
    <tr bgcolor="#000099">
      <td width="35%"><div align="left"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></strong></div></td>
      <td width="13%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Criado por</font></strong></div></td>
      <td width="12%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif"> Cria&ccedil;&atilde;o</font></strong></div></td>
      <td width="12%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Vezes Exportadas </font></strong></div></td>
      <td width="18%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">&Uacute;ltima Exporta&ccedil;&atilde;o </font></strong></div></td>
      <td width="10%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Usu-Func</font></strong></div></td>
    </tr>
	<% do while not rds_Lote.Eof %>
    <tr>
      <td><div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="consulta_lote_usu_func.asp?str_Tipo_Saida=Tela&pLote=<%=rds_Lote("LOTE_NR_SEQ_LOTE")%>&pDescLote=<%=rds_Lote("LOTE_TX_DESCRICAO")%>&pVezesImp=<%=rds_Lote("LOTE_NR_QTD_EXPORTACAO")%>"><%=rds_Lote("LOTE_NR_SEQ_LOTE")%> - <%=rds_Lote("LOTE_TX_DESCRICAO")%></a></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("ATUA_CD_NR_USUARIO")%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=formatadata("MDA","DMA",rds_Lote("LOTE_DT_ENVIO"))%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("LOTE_NR_QTD_EXPORTACAO")%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("ATUA_DT_ATUALIZACAO")%></font></div></td>
      <td><div align="center"><a href="consulta_conteudo_lote.asp?str_Tipo_Saida=Tela&pLote=<%=rds_Lote("LOTE_NR_SEQ_LOTE")%>&pDescLote=<%=rds_Lote("LOTE_TX_DESCRICAO")%>&pVezesImp=<%=rds_Lote("LOTE_NR_QTD_EXPORTACAO")%>&pOrdem=1"><img src="../../imagens/b04.gif" width="16" height="16" border="0"></a></div></td>
    </tr>
    <tr>
      <td height="25" colspan="6"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;o:</font></strong> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("LOTE_TX_ORGAO_SELEC")%></font></td>
    </tr>
    <tr>
      <td height="25" colspan="6"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Fun&ccedil;&atilde;o</font></strong>: <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("LOTE_TX_FUNCAO_SELEC")%></font></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="1" colspan="6"></td>
    </tr>
	<% rds_Lote.movenext
	Loop %>
    <tr>
      <td height="25">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
<%
rds_Lote.close
set rds_Lote = Nothing
conn_Cogest.Close
set conn_Cogest = Nothing
%>
