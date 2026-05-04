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
<form name="frm1" method="post" action="gera_apoio_modulo.asp">
<table width="80%" border="0">
  <tr>
    <td width="72%"><div align="center"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>CONSULTA 
          POR ORG&Atilde;O APOIADO x ASSUNTO</strong></font></div></td>
      <td width="28%">&nbsp;</td>
  </tr>
</table>
  <table width="70%" height="315" border="0">
    <tr> 
      <td width="13%">&nbsp;</td>
      <td width="25" colspan="2">&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr> 
      <td height="38">&nbsp;</td>
      <td colspan="2"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;o 
        Aglutinador</font></strong></td>
      <td colspan="4"><select name="selAglu" id="selAglu">
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
      <td height="41">&nbsp;</td>
      <td colspan="2"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Atribui&ccedil;&atilde;o</font></strong></td>
      <td colspan="4"><select name="selAtrib" id="selAtrib">
          <option value="1">APOIADOR LOCAL</option>
          <option value="2">MULTIPLICADOR</option>
        </select></td>
    </tr>
    <tr> 
      <td height="32">&nbsp;</td>
      <td colspan="2"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Classifica&ccedil;&atilde;o</font></strong></td>
      <td width="7%"><p align="right"> 
          <label> 
          <input name="selClass" type="radio" value="1" checked>
          </label>
        </p></td>
      <td width="18%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Quantitativo</font></strong></td>
      <td width="6%"><div align="right"><font color="#000066"> 
          <input type="radio" name="selClass" value="2">
          </font></div></td>
      <td width="31%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Nominal</font></strong></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right"></div></td>
      <td><div align="right"><img src="../../../imagens/confirma_f02.gif" width="24" height="24" onClick="Confirma()"></div></td>
      <td colspan="4"> <strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        Montar Relat&oacute;rio</font></strong></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td colspan="4">&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<p>&nbsp;</p><p>&nbsp;</p></body>
</html>
