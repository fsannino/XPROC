<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE")
%>
<html>
<head>
<script>
function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href.href='altera_Atividade1.asp?selAtiv='+document.frm1.selAtividade.value
	 }
 }
</SCRIPT>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="">
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
      <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Alteração 
        de Atividade</font></td>
      <td width="26%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2">&nbsp;Selecione 
        a Atividade que deseja alterar</font></b></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <p> 
          <select size="1" name="selAtividade" onchange="javascript:Confirma()">
            <option value="0">== Selecione a Atividade ==</option>
            <%do while not rs.EOF %>
            <option value=<%=rs("ATCA_CD_ATIVIDADE_CARGA")%>><%=rs("ATCA_TX_DESC_ATIVIDADE")%></option>
            <% rs.movenext
  Loop
  %>
          </select>
        </p>
        </font></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="12%">&nbsp;</td>
      <td width="63%">&nbsp;</td>
    </tr>
  </table> 
</form>
</body>
</html>
