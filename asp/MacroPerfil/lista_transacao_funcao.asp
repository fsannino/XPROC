<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
if (Request("selFuncao") <> "") then 
    str_Funcao = Request("selFuncao")
else
    str_Funcao = "0"
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQL = str_SQL & " SELECT "    
str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, "
str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO.FUNE_TX_DESC_FUNCAO_NEGOCIO, dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO, "
str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO.FUNE_TX_TP_FUN_NEG, dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO_PAI, "
str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO.FUNE_TX_INDICA_REFERENCIADA, "
str_SQL = str_SQL & " dbo.MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL = str_SQL & " FROM dbo.FUNCAO_NEGOCIO INNER JOIN"
str_SQL = str_SQL & " dbo.MEGA_PROCESSO ON "
str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
'str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO LEFT OUTER JOIN"
'str_SQL = str_SQL & " dbo.SUB_MODULO ON dbo.FUNCAO_NEGOCIO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA"
str_SQL = str_SQL & " WHERE  dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'"
str_SQL = str_SQL & " order by dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "

'response.Write(str_SQL)
Set rdsFuncao = Conn_db.Execute(str_SQL)
if not rdsFuncao.EOF then
   str_CdFuncao = rdsFuncao("FUNE_CD_FUNCAO_NEGOCIO")
   str_Titulo = rdsFuncao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
   str_Desc = rdsFuncao("FUNE_TX_DESC_FUNCAO_NEGOCIO")
end if
rdsFuncao.close
set rdsFuncao = Nothing

str_SQL = ""
str_SQL = str_SQL & " SELECT DISTINCT "
str_SQL = str_SQL & " dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO, dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL = str_SQL & " dbo.TRANSACAO.TRAN_TX_DESC_TRANSACAO"
str_SQL = str_SQL & " FROM  dbo.FUN_NEG_TRANSACAO INNER JOIN"
str_SQL = str_SQL & " dbo.TRANSACAO ON dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = dbo.TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL = str_SQL & " WHERE dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_CdFuncao & "'"
Set rdsFuncaoTrans = Conn_db.Execute(str_SQL)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Neg&oacute;cio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="650" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#003399"> 
    <td width="55">&nbsp;</td>
    <td width="561">&nbsp;</td>
    <td width="77">&nbsp;</td>
  </tr>
  <tr bgcolor="#003399"> 
    <td>&nbsp;</td>
    <td> 
      <div align="center"><font color="#FFFFFF" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Fun&ccedil;&atilde;o 
        R/3</strong></font></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr bgcolor="#003399"> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="650" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="76">&nbsp;</td>
    <td width="617"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:close()">Fechar 
        janela</a></font></div></td>
  </tr>
  <tr> 
    <td><div align="right"><font color="#006699" size="2" face="Verdana, Arial, Helvetica, sans-serif">Fun&ccedil;&atilde;o:</font></div></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_CdFuncao%></font></td>
  </tr>
  <tr> 
    <td><div align="right"><font color="#006699" size="2" face="Verdana, Arial, Helvetica, sans-serif">T&iacute;tulo:</font></div></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Titulo%></font></td>
  </tr>
  <tr> 
    <td valign="top">
<div align="right"><font color="#006699" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o:</font></div></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Desc%></font></td>
  </tr>
  <tr> 
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="650" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="27">&nbsp;</td>
    <td width="666"><strong><font color="#006699" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&otilde;es</font></strong></td>
  </tr>
  <tr> 
    <td height="5">&nbsp;</td>
    <td height="5" bgcolor="#0066FF">&nbsp;</td>
  </tr>
  <% If not rdsFuncaoTrans.EOF then 
     do while not rdsFuncaoTrans.EOF
	 
	    str_SQL = ""
		str_SQL = str_SQL & " SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
		str_SQL = str_SQL & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
        str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO_MEGA INNER JOIN"
        str_SQL = str_SQL & " " & Session("PREFIXO") & "MEGA_PROCESSO ON "
        str_SQL = str_SQL & " " & Session("PREFIXO") & "TRANSACAO_MEGA.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
        str_SQL = str_SQL & " WHERE " & Session("PREFIXO") & "TRANSACAO_MEGA.TRAN_CD_TRANSACAO = '" & rdsFuncaoTrans("TRAN_CD_TRANSACAO") & "'" 
		str_SQL = str_SQL & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "				   
		Set rdsExiste2 = Conn_db.Execute(str_SQL)				   
		loo_Existe = False
		str_Mega = "         - Dono : "
		a = ""
				   IF not rdsExiste2.EOF then
				      Do While not rdsExiste2.EOF
					     str_Mega = str_Mega & rdsExiste2("MEPR_TX_DESC_MEGA_PROCESSO") & " / "
					     rdsExiste2.Movenext
				      Loop
				   else
				   	  str_Mega = str_Mega & "   -  em processo de definição de dono "
				   end if
				   rdsExiste2.close
				   set rdsExiste2 = Nothing
	

  %>
  <tr> 
    <td><div align="right"></div></td>
    <td><font color="#0033CC" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsFuncaoTrans("TRAN_CD_TRANSACAO")%> - <strong><%=rdsFuncaoTrans("TRAN_TX_DESC_TRANSACAO")%></strong></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><i><font color="#999999" size="1"><%=str_Mega%> </font></i></font></td>
  </tr>
  <% rdsFuncaoTrans.movenext
     Loop  
  else %>
  <tr> 
    <td>&nbsp;</td>
    <td><div align="center"><font color="#003399" size="3" face="Verdana, Arial, Helvetica, sans-serif">N&atilde;o 
        encontrado transa&ccedil;&otilde;es para esta Fun&ccedil;&atilde;o</font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <% end if 
     rdsFuncaoTrans.close
     set rdsFuncaoTrans = Nothing
  %>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:close()">Fechar 
        janela</a></font></div></td>
  </tr>
</table>
</body>
</html>
