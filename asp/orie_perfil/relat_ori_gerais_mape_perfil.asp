<%@LANGUAGE="VBSCRIPT"%> 
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " ORIE_NR_SEQUENCIAL "
str_SQL = str_SQL & " , ORIE_TX_ORIENTACOES "
str_SQL = str_SQL & " , ORIE_NR_ORDENACAO "
str_SQL = str_SQL & " from PERFIL_ORIEN_GERAL "
'response.Write(str_SQL)
set rdsOrient = db.Execute(str_SQL)

str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " ORTE_NR_SEQUENCIAL "
str_SQL = str_SQL & " , ORTE_TX_TERMO "
str_SQL = str_SQL & " , ORTE_TX_DESCRICAO "
str_SQL = str_SQL & " from PERFIL_ORIEN_GERAL_TERMOS "
'response.Write(str_SQL)
set rdsTermo = db.Execute(str_SQL)


%>
<html>
<head>
<script>
function manda()
{
//alert('altera_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value)
//+'&selSubModulo='+
//alert(document.frm1.selSubModulo.value)
window.location.href='altera_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value
}

function Confirma()
{
   document.frm1.submit(); 
}
</script>
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
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="90%">&nbsp;</td>
      <td width="5%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="90%" height="40" bgcolor="#666666"> 
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Orienta&ccedil;&otilde;es 
          gerais ao Mapeamento de Perfil </strong></font></div></td>
      <td width="5%">&nbsp;</td>
    </tr>
  </table>
  
<table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td width="89%"></td>
  </tr>
  <% Do While not rdsOrient.EOF %>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsOrient("ORIE_NR_ORDENACAO")%> - <%=rdsOrient("ORIE_TX_ORIENTACOES")%></font></td>
  </tr>
  <%  rdsOrient.movenext
	Loop %>
</table> 
<table width="90%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
  <tr> 
    <td colspan="2" bgcolor="#999999"> <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><strong>Termos 
        Novos Relevantes</strong></font></div></td>
  </tr>
  <tr> 
    <td width="25%"><div align="center"><font color="#999999" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Termo</strong></font></div></td>
    <td width="65%"><div align="left"><font color="#999999" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>Descri&ccedil;&atilde;o</strong></font></div></td>
  </tr>
  <% do while not rdsTermo.EOF %>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsTermo("ORTE_TX_TERMO")%></font></td>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsTermo("ORTE_TX_DESCRICAO")%></font></td>
  </tr>
  <% rdsTermo.movenext
  Loop %>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
