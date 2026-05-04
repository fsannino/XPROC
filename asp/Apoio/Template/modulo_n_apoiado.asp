<!--#include file="../conn_consulta.asp" -->
<html>
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs1=db.execute("SELECT * FROM ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")

%>
<head>
<title>Base de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>
<script>
function Confirma()
{
if(document.frm1.selAglu.selectedIndex == 0)
{
	alert('Você deve selecionar um ÓRGÃO AGLUTINADOR');
	document.frm1.selAglu.focus();
	return;
}
else
{
	document.frm1.submit()
}
}
</script> 
<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="gera_modulo_n_apoiado.asp">
  <table width="67%" border="0">
    <tr>
    <td width="94%"><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>CONSULTA 
          DE ASSUNTOS N&Atilde;O APOIADOS POR &Oacute;RG&Atilde;O</strong></font></div></td>
      <td width="6%">&nbsp;</td>
  </tr>
</table>
  <table width="70%" height="215" border="0">
    <tr> 
      <td width="97" height="21">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td width="375">&nbsp;</td>
    </tr>
    <tr> 
      <td height="21">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="38">&nbsp;</td>
      <td colspan="2"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;o 
        Aglutinador</font></strong></td>
      <td><select name="selAglu" id="selAglu">
          <option value="0">== Selecione ==</option>
          <%
	  do until rs1.eof=true
	  %>
          <option value="<%=rs1("AGLU_CD_AGLUTINADO")%>"><%=rs1("AGLU_SG_AGLUTINADO")%></option>
          <%
	  rs1.movenext
	  loop
	  %>
        </select></td>
    </tr>
    <tr>
      <td height="21">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="26">&nbsp;</td>
      <td width="1"><div align="right"></div></td>
      <td width="215"><div align="right"><img src="../../../imagens/confirma_f02.gif" width="24" height="24" onClick="Confirma()"></div></td>
      <td> <strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        Montar Relat&oacute;rio</font></strong></td>
    </tr>
    <tr> 
      <td height="21">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="21">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p></body>
</html>
