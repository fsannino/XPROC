<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")

SELECT CASE REQUEST("ORDER")
	CASE 1
		VALOR="MODU_CD_MODULO"
	CASE 2
		VALOR="MODU_TX_DESC_MODULO"
	CASE ELSE
		VALOR="MODU_TX_DESC_MODULO"
END SELECT

db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 ORDER BY " & VALOR)
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
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
    <td width="24%">&nbsp;</td>
    <td width="50%">&nbsp;</td>
    <td width="26%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="24%">&nbsp;</td>
    <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      dos Agrupamento (Master List R/3) Cadastrados</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%"><font face="Verdana"><b>&nbsp;</b></font></td>
    <td width="63%" valign="bottom" align="left"><font size="1" face="Verdana"><b>Clique
      na coluna desejada para ordenar</b></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%" bgcolor="#0066CC"><b><a href="consulta_modulo.asp?order=1"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></a></b></td>
    <td width="63%" bgcolor="#0066CC"><b><a href="consulta_modulo.asp?order=2"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Agrupamento 
      (Master List R/3</font></a><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">)</font></b></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%do while not rs.EOF %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("MODU_CD_MODULO")%></font></td>
    <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("MODU_TX_DESC_MODULO")%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <% rs.movenext
  Loop
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
</html>
