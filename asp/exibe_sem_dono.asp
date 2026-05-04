<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

SERVER.SCRIPTTIMEOUT = 99999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

'set rs=db.execute("SELECT DISTINCT TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "RELACAO_FINAL ORDER BY TRAN_CD_TRANSACAO")

ls_SQl = ""
ls_SQl = ls_SQl & " SELECT distinct " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO, "
ls_SQl = ls_SQl & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
ls_SQl = ls_SQl & " FROM " & Session("PREFIXO") & "TRANSACAO INNER JOIN"
ls_SQl = ls_SQl & " " & Session("PREFIXO") & "RELACAO_FINAL ON "
ls_SQl = ls_SQl & " " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO"
ls_SQl = ls_SQl & " order by 1 "

set rs=db.execute(ls_SQl)
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
<%if request("excel")=0 then%>
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
          <td width="64%"><a href="exibe_sem_dono.asp?excel=1" target="_blank"><img border="0" src="../imagens/exp_excel.gif" width="78" height="29"></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%" bgcolor="white">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="62%">
      <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      de Transa&ccedil;&otilde;es sem Dono</font></p>
    </td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"></td>
    <td width="62%"></td>
    <td width="18%"></td>
  </tr>
  <tr> 
    <td width="20%"></td>
    <td width="62%"></td>
    <td width="18%"></td>
  </tr>
</table>
<table border="0" width="100%" height="68">
  <tr>
    <td width="100%" height="40"><font face="Verdana" size="2"><b>Transações</b></font></td>
  </tr>
  <%
  int_Contador = 0
  DO UNTIL RS.EOF=TRUE
  
  SET TMP=DB.EXECUTE("select * from " & Session("PREFIXO") & "transacao_mega where tran_cd_transacao='" & rs("tran_cd_transacao") & "'")

  if tmp.eof=true then
     int_Contador = int_Contador + 1
  if cor="white" then
  	cor="#D8D8D8"
  else
  	cor="white"
  end if
  
  %>
	  <tr bgcolor="<%=cor%>">
        <td height="16"><font face="Verdana" size="2"><%=rs("TRAN_CD_TRANSACAO")%>-<%=rs("TRAN_TX_DESC_TRANSACAO")%></font>
  </tr>
  <%
  end if
  RS.MOVENEXT
  LOOP
  %>
</table>
<p><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>Total de registros 
  : <%=int_Contador%></b></font></p>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

SERVER.SCRIPTTIMEOUT = 99999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

'set rs=db.execute("SELECT DISTINCT TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "RELACAO_FINAL ORDER BY TRAN_CD_TRANSACAO")

ls_SQl = ""
ls_SQl = ls_SQl & " SELECT distinct " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO, "
ls_SQl = ls_SQl & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
ls_SQl = ls_SQl & " FROM " & Session("PREFIXO") & "TRANSACAO INNER JOIN"
ls_SQl = ls_SQl & " " & Session("PREFIXO") & "RELACAO_FINAL ON "
ls_SQl = ls_SQl & " " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO"
ls_SQl = ls_SQl & " order by 1 "

set rs=db.execute(ls_SQl)
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
<%if request("excel")=0 then%>
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
          <td width="64%"><a href="exibe_sem_dono.asp?excel=1" target="_blank"><img border="0" src="../imagens/exp_excel.gif" width="78" height="29"></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%" bgcolor="white">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="62%">
      <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      de Transa&ccedil;&otilde;es sem Dono</font></p>
    </td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"></td>
    <td width="62%"></td>
    <td width="18%"></td>
  </tr>
  <tr> 
    <td width="20%"></td>
    <td width="62%"></td>
    <td width="18%"></td>
  </tr>
</table>
<table border="0" width="100%" height="68">
  <tr>
    <td width="100%" height="40"><font face="Verdana" size="2"><b>Transações</b></font></td>
  </tr>
  <%
  int_Contador = 0
  DO UNTIL RS.EOF=TRUE
  
  SET TMP=DB.EXECUTE("select * from " & Session("PREFIXO") & "transacao_mega where tran_cd_transacao='" & rs("tran_cd_transacao") & "'")

  if tmp.eof=true then
     int_Contador = int_Contador + 1
  if cor="white" then
  	cor="#D8D8D8"
  else
  	cor="white"
  end if
  
  %>
	  <tr bgcolor="<%=cor%>">
        <td height="16"><font face="Verdana" size="2"><%=rs("TRAN_CD_TRANSACAO")%>-<%=rs("TRAN_TX_DESC_TRANSACAO")%></font>
  </tr>
  <%
  end if
  RS.MOVENEXT
  LOOP
  %>
</table>
<p><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>Total de registros 
  : <%=int_Contador%></b></font></p>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
