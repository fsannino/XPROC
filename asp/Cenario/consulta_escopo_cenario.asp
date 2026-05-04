<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../../asp/protege/protege.asp" -->
<%
Session.LCID=1046

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

ssql=""
ssql="SELECT DISTINCT CENA_CD_CENARIO, ATUA_DT_ATUALIZACAO FROM CENARIO_VALIDACAO "
ssql=ssql+ "WHERE CEVA_TX_SITUACAO = 1 "
ssql=ssql+ "ORDER BY CENA_CD_CENARIO, ATUA_DT_ATUALIZACAO" 

set rs=conn_db.execute(ssql)
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
</script>

<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="valida_altera_escopo.asp">
  <input type="hidden" name="INC" size="20" value="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><input type="hidden" name="Atual" size="9" value="<%=valor_atual%>"></font>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0" width="19" height="20"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26">&nbsp;</td>
          <td width="50">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="195">&nbsp;</td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  &nbsp;<p><b><font face="Verdana" color="#00509F">Consulta Validação de Escopo de Cenário</font></b></p>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="392" id="AutoNumber1" height="45">
             <tr>
                        <td width="170" height="18" bgcolor="#000080"><b><font face="Verdana" color="#FFFFFF" size="2">Cenário</font></b></td>
                        <td width="219" height="18" bgcolor="#000080"><b><font face="Verdana" color="#FFFFFF" size="2">Data de Entrada no Escopo</font></b></td>
             </tr>
             <%
             DO UNTIL RS.EOF=TRUE
             %>
             <tr>
                        <td width="170" height="26"><b><font face="Verdana" color="#000080" size="1"><%=RS("CENA_CD_CENARIO")%></font></b></td>
                        <td width="219" height="26"><b><font face="Verdana" color="#000080" size="1"><%=RS("ATUA_DT_ATUALIZACAO")%></font></b></td>
             </tr>
             <%
             RS.MOVENEXT
             LOOP
             %>
  </table>
  </form>
<p>&nbsp;</p>
</body>

</html>