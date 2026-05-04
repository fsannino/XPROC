<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_Trans = " SELECT COUNT(MEPR_CD_MEGA_PROCESSO) AS Expr1 "
str_SQL_Trans = str_SQL_Trans & " ,TRAN_CD_TRANSACAO "
str_SQL_Trans = str_SQL_Trans & " FROM " & Session("PREFIXO") & "RELACAO_FINAL "
str_SQL_Trans = str_SQL_Trans & " GROUP BY TRAN_CD_TRANSACAO "
str_SQL_Trans = str_SQL_Trans & " HAVING (COUNT(MEPR_CD_MEGA_PROCESSO) > 1)"
str_SQL_Trans = str_SQL_Trans & " ORDER BY  TRAN_CD_TRANSACAO "

contador = 0
'response.write str_SQL_Trans

Set rdsTransacao= Conn_db.Execute(str_SQL_Trans)

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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../index.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="17%">&nbsp;</td>
    <td width="69%">&nbsp;</td>
    <td width="14%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="17%">&nbsp;</td>
    <td width="69%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      de Transa&ccedil;&otilde;es em mais de um Mega-Processo - old2</font></td>
    <td width="14%">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <% If not rdsTransacao.EOF then 
	      Do While not rdsTransacao.EOF
		     str_SQL_Mega = " SELECT DISTINCT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO  "
             str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO "
             str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "			 
             str_SQL_Mega = str_SQL_Mega & " FROM " & Session("PREFIXO") & "RELACAO_FINAL, " & Session("PREFIXO") & "MEGA_PROCESSO, " & Session("PREFIXO") & "TRANSACAO  "
             str_SQL_Mega = str_SQL_Mega & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'"
			' str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & "11 "						 			 			 
             str_SQL_Mega = str_SQL_Mega & " and " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
			 str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "			 
             str_SQL_Mega = str_SQL_Mega & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
			 'response.write str_SQL_Mega
             Set rdsMegaTot= Conn_db.Execute(str_SQL_Mega)
			 if not rdsMegaTot.EOF then
		        str_SQL_Mega = " SELECT DISTINCT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
                str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO "
                str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "			 
                str_SQL_Mega = str_SQL_Mega & " FROM " & Session("PREFIXO") & "RELACAO_FINAL, " & Session("PREFIXO") & "MEGA_PROCESSO, " & Session("PREFIXO") & "TRANSACAO  "
                str_SQL_Mega = str_SQL_Mega & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'"
                str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO <> " & rdsMegaTot("MEPR_CD_MEGA_PROCESSO") 
			   ' str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & "11 "						 
                str_SQL_Mega = str_SQL_Mega & " and " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
			 str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "			 
             str_SQL_Mega = str_SQL_Mega & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
			 'response.write str_SQL_Mega
             Set rdsMegaRepete= Conn_db.Execute(str_SQL_Mega)

			 if not rdsMegaRepete.EOF then
			 contador = contador + 1
	%>
    <td width="9%">&nbsp;</td>
    <td width="12%" bgcolor="#0066CC" style="color: #FFFFFF"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o</font></b></td>
    <td width="63%" bgcolor="#0066CC" style="color: #FFFFFF"><b></b></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsMegaRepete("TRAN_CD_TRANSACAO")%> - <%=rdsMegaRepete("TRAN_TX_DESC_TRANSACAO")%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#0066CC">Mega-Processo</font></b></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <% 'rdsMegaTot.movefirst
  Do While not rdsMegaTot.EOF %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsMegaTot("MEPR_CD_MEGA_PROCESSO")%>- <%=rdsMegaTot("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <% rdsMegaTot.movenext
       Loop
    end if 
	   rdsMegaRepete.Close
	 end if  
	   rdsMegaTot.Close
     rdsTransacao.movenext
      Loop
	   rdsTransacao.Close
    
	 else %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">N&atilde;o 
      existem Transa&ccedil;&otilde;es em mais de um Mega-Processo</font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <% end if %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr>
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">Total impresso :<%=contador%> </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr>
    <td width="9%" height="2">&nbsp;</td>
    <td width="12%" height="2">&nbsp;</td>
    <td width="63%" height="2">&nbsp;</td>
    <td width="16%" height="2">&nbsp;</td>
  </tr>
</table>
<p>OK</p>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_Trans = " SELECT COUNT(MEPR_CD_MEGA_PROCESSO) AS Expr1 "
str_SQL_Trans = str_SQL_Trans & " ,TRAN_CD_TRANSACAO "
str_SQL_Trans = str_SQL_Trans & " FROM " & Session("PREFIXO") & "RELACAO_FINAL "
str_SQL_Trans = str_SQL_Trans & " GROUP BY TRAN_CD_TRANSACAO "
str_SQL_Trans = str_SQL_Trans & " HAVING (COUNT(MEPR_CD_MEGA_PROCESSO) > 1)"
str_SQL_Trans = str_SQL_Trans & " ORDER BY  TRAN_CD_TRANSACAO "

contador = 0
'response.write str_SQL_Trans

Set rdsTransacao= Conn_db.Execute(str_SQL_Trans)

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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../index.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="17%">&nbsp;</td>
    <td width="69%">&nbsp;</td>
    <td width="14%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="17%">&nbsp;</td>
    <td width="69%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      de Transa&ccedil;&otilde;es em mais de um Mega-Processo - old2</font></td>
    <td width="14%">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <% If not rdsTransacao.EOF then 
	      Do While not rdsTransacao.EOF
		     str_SQL_Mega = " SELECT DISTINCT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO  "
             str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO "
             str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "			 
             str_SQL_Mega = str_SQL_Mega & " FROM " & Session("PREFIXO") & "RELACAO_FINAL, " & Session("PREFIXO") & "MEGA_PROCESSO, " & Session("PREFIXO") & "TRANSACAO  "
             str_SQL_Mega = str_SQL_Mega & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'"
			' str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & "11 "						 			 			 
             str_SQL_Mega = str_SQL_Mega & " and " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
			 str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "			 
             str_SQL_Mega = str_SQL_Mega & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
			 'response.write str_SQL_Mega
             Set rdsMegaTot= Conn_db.Execute(str_SQL_Mega)
			 if not rdsMegaTot.EOF then
		        str_SQL_Mega = " SELECT DISTINCT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
                str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO "
                str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "			 
                str_SQL_Mega = str_SQL_Mega & " FROM " & Session("PREFIXO") & "RELACAO_FINAL, " & Session("PREFIXO") & "MEGA_PROCESSO, " & Session("PREFIXO") & "TRANSACAO  "
                str_SQL_Mega = str_SQL_Mega & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'"
                str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO <> " & rdsMegaTot("MEPR_CD_MEGA_PROCESSO") 
			   ' str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & "11 "						 
                str_SQL_Mega = str_SQL_Mega & " and " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
			 str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "			 
             str_SQL_Mega = str_SQL_Mega & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
			 'response.write str_SQL_Mega
             Set rdsMegaRepete= Conn_db.Execute(str_SQL_Mega)

			 if not rdsMegaRepete.EOF then
			 contador = contador + 1
	%>
    <td width="9%">&nbsp;</td>
    <td width="12%" bgcolor="#0066CC" style="color: #FFFFFF"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o</font></b></td>
    <td width="63%" bgcolor="#0066CC" style="color: #FFFFFF"><b></b></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsMegaRepete("TRAN_CD_TRANSACAO")%> - <%=rdsMegaRepete("TRAN_TX_DESC_TRANSACAO")%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#0066CC">Mega-Processo</font></b></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <% 'rdsMegaTot.movefirst
  Do While not rdsMegaTot.EOF %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsMegaTot("MEPR_CD_MEGA_PROCESSO")%>- <%=rdsMegaTot("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <% rdsMegaTot.movenext
       Loop
    end if 
	   rdsMegaRepete.Close
	 end if  
	   rdsMegaTot.Close
     rdsTransacao.movenext
      Loop
	   rdsTransacao.Close
    
	 else %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">N&atilde;o 
      existem Transa&ccedil;&otilde;es em mais de um Mega-Processo</font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <% end if %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr>
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">Total impresso :<%=contador%> </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr>
    <td width="9%" height="2">&nbsp;</td>
    <td width="12%" height="2">&nbsp;</td>
    <td width="63%" height="2">&nbsp;</td>
    <td width="16%" height="2">&nbsp;</td>
  </tr>
</table>
<p>OK</p>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
