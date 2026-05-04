<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")

select case request("ORDER")
	CASE 1
		VALOR="TRAN_CD_TRANSACAO"
	CASE 2
		VALOR="TRAN_TX_DESC_TRANSACAO"
	CASE ELSE
		VALOR="TRAN_TX_DESC_TRANSACAO"
END select

db.Open Session("Conn_String_Cogest_Gravacao")

db.cursorlocation = 3
set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY " & VALOR)
i = 0
reg = rs.recordcount
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

<script>
function exibe_dono(trans)
{
var a=trans;
window.open("mostra_dono.asp?transacao=" + a + "","_blank","width=330,height=200,history=0,scrollbars=0,titlebar=0,resizable=0")
}
</script>

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
    <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      das Transações Cadastradas</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%"><b><font size="1" face="Verdana">Clique na coluna desejada
      para ordenar</font></b></td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%" bgcolor="#0066CC"><b><a href="consulta_trans.asp?order=1"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></a></b></td>
    <td width="62%" bgcolor="#0066CC"><b><a href="consulta_trans.asp?order=2"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o</font></a></b></td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="6%" align="center"></td>
    <td width="10%" align="center"></td>
  </tr>
  <%'do while not rs.EOF 
     do until i = reg
  if cor="white" then
  	cor="#DADADA"
  else
  	cor="white"
  end if  
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%" bgcolor=<%=cor%>><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("TRAN_CD_TRANSACAO")%></font></td>
    <td width="62%" bgcolor=<%=cor%>><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("TRAN_TX_DESC_TRANSACAO")%></font></td>
    <td width="6%" align="center" bgcolor=<%=cor%>><a href="javascript:exibe_dono('<%=rs("TRAN_CD_TRANSACAO")%>')"><img border="0" src="../imagens/b04.gif" alt="Clique aqui para saber o Dono desta Transação"></a>&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <% i = i + 1
  rs.movenext
  Loop
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
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

select case request("ORDER")
	CASE 1
		VALOR="TRAN_CD_TRANSACAO"
	CASE 2
		VALOR="TRAN_TX_DESC_TRANSACAO"
	CASE ELSE
		VALOR="TRAN_TX_DESC_TRANSACAO"
END select

db.Open Session("Conn_String_Cogest_Gravacao")

db.cursorlocation = 3
set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY " & VALOR)
i = 0
reg = rs.recordcount
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

<script>
function exibe_dono(trans)
{
var a=trans;
window.open("mostra_dono.asp?transacao=" + a + "","_blank","width=330,height=200,history=0,scrollbars=0,titlebar=0,resizable=0")
}
</script>

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
    <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      das Transações Cadastradas</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%"><b><font size="1" face="Verdana">Clique na coluna desejada
      para ordenar</font></b></td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%" bgcolor="#0066CC"><b><a href="consulta_trans.asp?order=1"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></a></b></td>
    <td width="62%" bgcolor="#0066CC"><b><a href="consulta_trans.asp?order=2"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o</font></a></b></td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="6%" align="center"></td>
    <td width="10%" align="center"></td>
  </tr>
  <%'do while not rs.EOF 
     do until i = reg
  if cor="white" then
  	cor="#DADADA"
  else
  	cor="white"
  end if  
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%" bgcolor=<%=cor%>><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("TRAN_CD_TRANSACAO")%></font></td>
    <td width="62%" bgcolor=<%=cor%>><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("TRAN_TX_DESC_TRANSACAO")%></font></td>
    <td width="6%" align="center" bgcolor=<%=cor%>><a href="javascript:exibe_dono('<%=rs("TRAN_CD_TRANSACAO")%>')"><img border="0" src="../imagens/b04.gif" alt="Clique aqui para saber o Dono desta Transação"></a>&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <% i = i + 1
  rs.movenext
  Loop
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="6%" align="center">&nbsp;</td>
    <td width="10%" align="center"></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
