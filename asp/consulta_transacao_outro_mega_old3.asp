<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<%
DIM vStr_Mega(10)

'SERVER.SCRIPTTIMEOUT = 99999999
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_Trans = " SELECT "
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_SQL_Trans = str_SQL_Trans & " FROM " & Session("PREFIXO") & "RELACAO_FINAL INNER JOIN "
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "TRANSACAO ON " 
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "
str_SQL_Trans = str_SQL_Trans & " WHERE (" & Session("PREFIXO") & "TRANSACAO.MEPR_CD_MEGA_PROCESSO IS NULL) "
str_SQL_Trans = str_SQL_Trans & " GROUP BY " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "

'str_SQL_Trans = " SELECT "
'str_SQL_Trans = str_SQL_Trans & " TRAN_CD_TRANSACAO "
'str_SQL_Trans = str_SQL_Trans & " , TRAN_TX_DESC_TRANSACAO "
'str_SQL_Trans = str_SQL_Trans & " , MEPR_CD_MEGA_PROCESSO "
'str_SQL_Trans = str_SQL_Trans & " FROM " & Session("PREFIXO") & "TRANSACAO "
'str_SQL_Trans = str_SQL_Trans & " WHERE MEPR_CD_MEGA_PROCESSO is null "

contador = 0

'response.write str_SQL_Trans

'Set rdsTransacao= Conn_db.Execute(str_SQL_Trans)

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
<!--#INCLUDE file="ADOVBS.INC" -->
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
    <td colspan="3" height="20">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="36%">&nbsp;</td>
          <td width="64%"><a href="consulta_transacao_outro_mega_excel.asp" target="_blank"><img src="../imagens/exp_excel.gif" width="78" height="29" border="0"></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      de Transa&ccedil;&otilde;es em mais de um Mega-Processo - </font></td>
    <td width="18%">&nbsp;</td>
  </tr>
</table>
<table border="0" cellpadding="0" width="1064">
  <tr>
    <td width="53" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="552" bgcolor="#FFFFFF"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o</font></td>
    <td width="42" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">1-SUP</font></b></div>
    </td>
    <td width="52" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">2-MES</font></b></div>
    </td>
    <td width="44" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">3-VEN</font></b></div>
    </td>
    <td width="46" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">4-EMP</font></b></div>
    </td>
    <td width="45" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">5-MAN</font></b></div>
    </td>
    <td width="43" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">6-PO&Ccedil;</font></b></div>
    </td>
    <td width="40" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">7-PRO</font></b></div>
    </td>
    <td width="47" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">8-QUA</font></b></div>
    </td>
    <td width="43" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">9-LOG</font></b></div>
    </td>
    <td width="44" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">10-PLA</font></b></div>
    </td>
    <td width="40" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">11-FIN</font></b></div>
    </td>
    <td width="71" bgcolor="#FFCC00"> 
      <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">DONO</font></b></div>
    </td>
  </tr>
  <%'************************************************************************%>
  <% Set rdsTransacao = Conn_db.Execute(str_SQL_Trans)

     set rdsMegaTot=Server.CreateObject("adodb.Recordset")

If not rdsTransacao.EOF Then
Do while not rdsTransacao.EOF

   str_SQL_Mega = " SELECT DISTINCT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO  "
   str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO "
   str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "			 
   str_SQL_Mega = str_SQL_Mega & " FROM " & Session("PREFIXO") & "RELACAO_FINAL, " & Session("PREFIXO") & "MEGA_PROCESSO, " & Session("PREFIXO") & "TRANSACAO  "
   str_SQL_Mega = str_SQL_Mega & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'"
   'str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & "2 "						 			 			 
   str_SQL_Mega = str_SQL_Mega & " and " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
   str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "			 
   str_SQL_Mega = str_SQL_Mega & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
   'set rdsMegaTot=Server.CreateObject("adodb.Recordset")
   rdsMegaTot.open str_SQL_Mega, Conn_db, adopenstatic
   'howmanyrecs=rdsMegaTot.recordcount
   if rdsMegaTot.recordcount > 1 then
      
     for  int_Index = 0 to 10
	    vStr_Mega(int_Index) = 0
     next		
      int_Index = 0
      do while not rdsMegaTot.EOF	     
         vStr_Mega(int_Index) = rdsMegaTot("MEPR_CD_MEGA_PROCESSO")
		 int_Index = int_Index + 1
	     rdsMegaTot.movenext
	  loop
	  contador = contador + 1
	  'for  int_Index = 0 to 10
	  '  response.write  vStr_Mega(int_Index)
      'next		
      if str_Color = "#D2D2D2" then
	     str_Color = "#FFFFFF"
	  else
	     str_Color = "#D2D2D2"
	  end if	 
    %>
  <tr bgcolor="<%=str_Color%>">
    <td width="53"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#333333" size="1"><%=rdsTransacao("TRAN_CD_TRANSACAO")%></font></b></font></td>
    <% str_SQL = ""
	   str_SQL = str_SQL & " SELECT "
       str_SQL = str_SQL & " " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO, "
       str_SQL = str_SQL & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "
       str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO "
       str_SQL = str_SQL & " WHERE " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'"	   
	   Set rdsDescTransacao= Conn_db.Execute(str_SQL)
	%>
    <td width="552"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b></b><font color="#333333" size="1">-<%=rdsDescTransacao("TRAN_TX_DESC_TRANSACAO")%></font></font></td>
    <td width="42"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 1 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
      </div>
      <div align="center"><b><%=str_Marca%></b></div>
    </td>
    <td width="52"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 2 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="44"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 3 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="46"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 4 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="45"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 5 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="43"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 6 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="40"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 7 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="47"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 8 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="43"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 9 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="44"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 10 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="40"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 11 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="71">&nbsp;</td>
  </tr>
  <%   rdsDescTransacao.close
  end if 
     rdsMegaTot.close
     rdsTransacao.movenext
   Loop %>
  <% else %>
</table>   
<p>N&atilde;o possui transa&ccedil;&otilde;es em mais de um mega.</p>
  <% end if 
  rdsTransacao.Close
  set rdsTransacao = Nothing
  %>
<p><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Total impressos 
  : <%=contador%></font></p>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<%
DIM vStr_Mega(10)

'SERVER.SCRIPTTIMEOUT = 99999999
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_Trans = " SELECT "
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_SQL_Trans = str_SQL_Trans & " FROM " & Session("PREFIXO") & "RELACAO_FINAL INNER JOIN "
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "TRANSACAO ON " 
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "
str_SQL_Trans = str_SQL_Trans & " WHERE (" & Session("PREFIXO") & "TRANSACAO.MEPR_CD_MEGA_PROCESSO IS NULL) "
str_SQL_Trans = str_SQL_Trans & " GROUP BY " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Trans = str_SQL_Trans & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "

'str_SQL_Trans = " SELECT "
'str_SQL_Trans = str_SQL_Trans & " TRAN_CD_TRANSACAO "
'str_SQL_Trans = str_SQL_Trans & " , TRAN_TX_DESC_TRANSACAO "
'str_SQL_Trans = str_SQL_Trans & " , MEPR_CD_MEGA_PROCESSO "
'str_SQL_Trans = str_SQL_Trans & " FROM " & Session("PREFIXO") & "TRANSACAO "
'str_SQL_Trans = str_SQL_Trans & " WHERE MEPR_CD_MEGA_PROCESSO is null "

contador = 0

'response.write str_SQL_Trans

'Set rdsTransacao= Conn_db.Execute(str_SQL_Trans)

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
<!--#INCLUDE file="ADOVBS.INC" -->
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
    <td colspan="3" height="20">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="36%">&nbsp;</td>
          <td width="64%"><a href="consulta_transacao_outro_mega_excel.asp" target="_blank"><img src="../imagens/exp_excel.gif" width="78" height="29" border="0"></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      de Transa&ccedil;&otilde;es em mais de um Mega-Processo - </font></td>
    <td width="18%">&nbsp;</td>
  </tr>
</table>
<table border="0" cellpadding="0" width="1064">
  <tr>
    <td width="53" bgcolor="#FFFFFF">&nbsp;</td>
    <td width="552" bgcolor="#FFFFFF"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o</font></td>
    <td width="42" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">1-SUP</font></b></div>
    </td>
    <td width="52" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">2-MES</font></b></div>
    </td>
    <td width="44" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">3-VEN</font></b></div>
    </td>
    <td width="46" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">4-EMP</font></b></div>
    </td>
    <td width="45" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">5-MAN</font></b></div>
    </td>
    <td width="43" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">6-PO&Ccedil;</font></b></div>
    </td>
    <td width="40" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">7-PRO</font></b></div>
    </td>
    <td width="47" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">8-QUA</font></b></div>
    </td>
    <td width="43" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">9-LOG</font></b></div>
    </td>
    <td width="44" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">10-PLA</font></b></div>
    </td>
    <td width="40" bgcolor="#FFFFCC"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">11-FIN</font></b></div>
    </td>
    <td width="71" bgcolor="#FFCC00"> 
      <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">DONO</font></b></div>
    </td>
  </tr>
  <%'************************************************************************%>
  <% Set rdsTransacao = Conn_db.Execute(str_SQL_Trans)

     set rdsMegaTot=Server.CreateObject("adodb.Recordset")

If not rdsTransacao.EOF Then
Do while not rdsTransacao.EOF

   str_SQL_Mega = " SELECT DISTINCT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO  "
   str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO "
   str_SQL_Mega = str_SQL_Mega & " ," & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "			 
   str_SQL_Mega = str_SQL_Mega & " FROM " & Session("PREFIXO") & "RELACAO_FINAL, " & Session("PREFIXO") & "MEGA_PROCESSO, " & Session("PREFIXO") & "TRANSACAO  "
   str_SQL_Mega = str_SQL_Mega & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'"
   'str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & "2 "						 			 			 
   str_SQL_Mega = str_SQL_Mega & " and " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
   str_SQL_Mega = str_SQL_Mega & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "			 
   str_SQL_Mega = str_SQL_Mega & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
   'set rdsMegaTot=Server.CreateObject("adodb.Recordset")
   rdsMegaTot.open str_SQL_Mega, Conn_db, adopenstatic
   'howmanyrecs=rdsMegaTot.recordcount
   if rdsMegaTot.recordcount > 1 then
      
     for  int_Index = 0 to 10
	    vStr_Mega(int_Index) = 0
     next		
      int_Index = 0
      do while not rdsMegaTot.EOF	     
         vStr_Mega(int_Index) = rdsMegaTot("MEPR_CD_MEGA_PROCESSO")
		 int_Index = int_Index + 1
	     rdsMegaTot.movenext
	  loop
	  contador = contador + 1
	  'for  int_Index = 0 to 10
	  '  response.write  vStr_Mega(int_Index)
      'next		
      if str_Color = "#D2D2D2" then
	     str_Color = "#FFFFFF"
	  else
	     str_Color = "#D2D2D2"
	  end if	 
    %>
  <tr bgcolor="<%=str_Color%>">
    <td width="53"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#333333" size="1"><%=rdsTransacao("TRAN_CD_TRANSACAO")%></font></b></font></td>
    <% str_SQL = ""
	   str_SQL = str_SQL & " SELECT "
       str_SQL = str_SQL & " " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO, "
       str_SQL = str_SQL & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "
       str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO "
       str_SQL = str_SQL & " WHERE " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'"	   
	   Set rdsDescTransacao= Conn_db.Execute(str_SQL)
	%>
    <td width="552"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b></b><font color="#333333" size="1">-<%=rdsDescTransacao("TRAN_TX_DESC_TRANSACAO")%></font></font></td>
    <td width="42"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 1 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
      </div>
      <div align="center"><b><%=str_Marca%></b></div>
    </td>
    <td width="52"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 2 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="44"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 3 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="46"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 4 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="45"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 5 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="43"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 6 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="40"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 7 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="47"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 8 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="43"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 9 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="44"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 10 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="40"> 
      <div align="center"> 
        <% int_Index = 0
	for  int_Index = 0 to 10
	    if vStr_Mega(int_Index) = 11 then
		   str_Marca = "x"
		   exit for
		else
		   str_Marca = ""
		end if
    next		
	%>
        <div align="center"><b><%=str_Marca%></b></div>
      </div>
    </td>
    <td width="71">&nbsp;</td>
  </tr>
  <%   rdsDescTransacao.close
  end if 
     rdsMegaTot.close
     rdsTransacao.movenext
   Loop %>
  <% else %>
</table>   
<p>N&atilde;o possui transa&ccedil;&otilde;es em mais de um mega.</p>
  <% end if 
  rdsTransacao.Close
  set rdsTransacao = Nothing
  %>
<p><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Total impressos 
  : <%=contador%></font></p>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
