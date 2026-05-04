<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")

SELECT CASE REQUEST("ORDER")
	CASE 1
		VALOR="USUA_CD_USUARIO"
	CASE 2
		VALOR="USUA_TX_NOME_USUARIO"		
	CASE 3
		VALOR="USUA_TX_CATEGORIA"
	CASE ELSE
		VALOR="USUA_TX_NOME_USUARIO"
END SELECT

db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "USUARIO ORDER BY " & VALOR)
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
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" border="0"></a>&nbsp;</div>
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
      de Usu&aacute;rio</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
</table>
<table width="697" border="0" cellspacing="0" cellpadding="0" height="77">
  <tr> 
    <td width="164" height="22">&nbsp;</td>
    <td width="112" height="22"><font face="Verdana"><b>&nbsp;</b></font></td>
    <td width="386" valign="bottom" align="left" height="22"><font size="1" face="Verdana"><b>Clique 
      na coluna desejada para ordenar</b></font></td>
    <td width="29" height="22">&nbsp;</td>
    <td width="29" height="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="164" height="21">&nbsp;</td>
    <td width="112" bgcolor="#0066CC" height="21"><b><a href="consulta_usuario.asp?order=1"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></a></b></td>
    <td width="386" bgcolor="#0066CC" height="21"><b><a href="consulta_usuario.asp?order=2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Nome</font></a></b></td>
    <td width="29" height="21" bgcolor="#0066CC"><b><a href="consulta_usuario.asp?order=3"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Categoria</font></a></b></td>
    <td width="29" height="21"></td>
  </tr>
  <%do while not rs.EOF %>
  <tr> 
    <td width="164" height="1">&nbsp;</td>
    <td width="112" height="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("USUA_CD_USUARIO")%></font></td>
    <td width="386" height="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("USUA_TX_NOME_USUARIO")%></font></td>
    <td width="29" height="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("USUA_TX_CATEGORIA")%></font></td>
    <td width="29" height="1"> 
      <p align="center"><a href="#" onclick=javascript:window.open("exibe_usu_mega.asp?ID=<%=rs("USUA_CD_USUARIO")%>","","width=420,height=280,status=0,toolbar=0,location=0,resizable=0")><img border="0" src="../imagens/icon_empresa.gif" alt="Exibe Mega-Processos relacionados..."></a></p>
    </td>
  </tr>
  <% rs.movenext
  Loop
  %>
  <tr> 
    <td width="164" height="19">&nbsp;</td>
    <td width="112" height="19">&nbsp;</td>
    <td width="386" height="19">&nbsp;</td>
    <td width="29" height="19">&nbsp;</td>
    <td width="29" height="19">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")

SELECT CASE REQUEST("ORDER")
	CASE 1
		VALOR="USUA_CD_USUARIO"
	CASE 2
		VALOR="USUA_TX_NOME_USUARIO"		
	CASE 3
		VALOR="USUA_TX_CATEGORIA"
	CASE ELSE
		VALOR="USUA_TX_NOME_USUARIO"
END SELECT

db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "USUARIO ORDER BY " & VALOR)
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
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" border="0"></a>&nbsp;</div>
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
      de Usu&aacute;rio</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
</table>
<table width="697" border="0" cellspacing="0" cellpadding="0" height="77">
  <tr> 
    <td width="164" height="22">&nbsp;</td>
    <td width="112" height="22"><font face="Verdana"><b>&nbsp;</b></font></td>
    <td width="386" valign="bottom" align="left" height="22"><font size="1" face="Verdana"><b>Clique 
      na coluna desejada para ordenar</b></font></td>
    <td width="29" height="22">&nbsp;</td>
    <td width="29" height="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="164" height="21">&nbsp;</td>
    <td width="112" bgcolor="#0066CC" height="21"><b><a href="consulta_usuario.asp?order=1"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></a></b></td>
    <td width="386" bgcolor="#0066CC" height="21"><b><a href="consulta_usuario.asp?order=2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Nome</font></a></b></td>
    <td width="29" height="21" bgcolor="#0066CC"><b><a href="consulta_usuario.asp?order=3"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Categoria</font></a></b></td>
    <td width="29" height="21"></td>
  </tr>
  <%do while not rs.EOF %>
  <tr> 
    <td width="164" height="1">&nbsp;</td>
    <td width="112" height="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("USUA_CD_USUARIO")%></font></td>
    <td width="386" height="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("USUA_TX_NOME_USUARIO")%></font></td>
    <td width="29" height="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("USUA_TX_CATEGORIA")%></font></td>
    <td width="29" height="1"> 
      <p align="center"><a href="#" onclick=javascript:window.open("exibe_usu_mega.asp?ID=<%=rs("USUA_CD_USUARIO")%>","","width=420,height=280,status=0,toolbar=0,location=0,resizable=0")><img border="0" src="../imagens/icon_empresa.gif" alt="Exibe Mega-Processos relacionados..."></a></p>
    </td>
  </tr>
  <% rs.movenext
  Loop
  %>
  <tr> 
    <td width="164" height="19">&nbsp;</td>
    <td width="112" height="19">&nbsp;</td>
    <td width="386" height="19">&nbsp;</td>
    <td width="29" height="19">&nbsp;</td>
    <td width="29" height="19">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
