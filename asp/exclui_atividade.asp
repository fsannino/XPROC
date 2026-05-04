<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")

db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE")
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
    <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Exclusão 
      de Atividade</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#800000">Clique 
      no link desejado para excluir</font></b></td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></b></td>
    <td width="63%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></b></td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="4%"></td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <%do while not rs.EOF %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("ATCA_CD_ATIVIDADE_CARGA")%></font></td>
    <td width="63%"><a href="exclusao.asp?tab=2&reg=<%=rs("ATCA_CD_ATIVIDADE_CARGA")%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("ATCA_TX_DESC_ATIVIDADE")%></font></a></td>
    <td width="4%"></td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <% rs.movenext
  Loop
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")

db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE")
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
    <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Exclusão 
      de Atividade</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#800000">Clique 
      no link desejado para excluir</font></b></td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></b></td>
    <td width="63%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></b></td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="4%"></td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <%do while not rs.EOF %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("ATCA_CD_ATIVIDADE_CARGA")%></font></td>
    <td width="63%"><a href="exclusao.asp?tab=2&reg=<%=rs("ATCA_CD_ATIVIDADE_CARGA")%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("ATCA_TX_DESC_ATIVIDADE")%></font></a></td>
    <td width="4%"></td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <% rs.movenext
  Loop
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
    <td width="63%">&nbsp;</td>
    <td width="4%">&nbsp;</td>
    <td width="5%"></td>
    <td width="7%"></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
