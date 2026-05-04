<!--#include file="conn_consulta.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs = db.execute("SELECT * FROM ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")

%>
<html>
<head>
<title>:: Consulta de Capacitação de Apoiadores Locais :::..</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function Consultar()
{
if(document.frm1.selOrgao.selectedIndex == 0)
{
alert("Você deve selecionar um ÓRGAO AGLUTINADOR!");
document.frm1.selOrgao.focus();
return;
}
else
{
document.frm1.submit()
}
}
</script>

<body bgcolor="#FFFFFF" text="#000000">
<p align="center">&nbsp;</p>
<p align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><b>Consulta 
  de Capacita&ccedil;&atilde;o de Apoiadores Locais</b></font></p>
<form name="frm1" method="post" action="gera_consulta.asp">
  <p align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2">Selecione 
    o &Oacute;rg&atilde;o Aglutinador</font></p>
  <p align="center"> <font face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
    <select name="selOrgao">
      <option value="0" selected>== Selecione ==</option>
	  <%
	  do until rs.eof=true
	  %>
      <option value="<%=rs("AGLU_CD_AGLUTINADO")%>"><%=rs("AGLU_SG_AGLUTINADO")%></option>	  
	  <%
	  rs.movenext
	  loop
	  %>
    </select>
    </font></p>
  <p align="center"> <font face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
    <input type="button" name="Submit" value="Consultar" onClick="Consultar()">
    </font></p>
</form>
</body>
</html>
