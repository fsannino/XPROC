<!--#include file="conn_consulta.asp" -->
<%
'======== ROTINA DO CONTADOR DE ACESSOS ==============

caminho = server.mappath("../../publico/demonstrativo/contador.txt")

set fs = server.CreateObject("Scripting.FileSystemObject")
set arquivo = fs.opentextfile(caminho)

visitas = arquivo.readline
visitas = visitas + 1

set arquivo = nothing
fs.deletefile(caminho)

set arquivo = fs.CreateTextFile(caminho)
linha = visitas
arquivo.writeline linha

set arquivo = nothing

'=================================================

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

set rs = db.execute("SELECT * FROM MEGA_PROCESSO WHERE MEPR_TX_ABREVIA NOT IN ('TI','GR') ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")
%>
<html>
<head>
	<title>:: Demostrativo de Cursos</title>
    <style type="text/css">
<!--
.style2 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
	color: #000066;
}
.style3 {
	font-size: x-small;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color: #000066;
}

.boton_box
{
	BORDER-RIGHT: black 1px solid;
	BORDER-TOP: black 1px solid;
	BORDER-COLOR: #000066;
	FONT-WEIGHT: bold;
	FONT-SIZE: 12px;
	WORD-SPACING: 2px;
	TEXT-TRANSFORM: capitalize;
	BORDER-LEFT: black 1px solid;
	COLOR: #000066;
	BORDER-BOTTOM: black 1px solid;
	FONT-STYLE: normal;
	FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif;
	BACKGROUND-COLOR: #FFFFFF;
}
-->
    </style>
</head>

<body>
<p align="center" class="style2">&nbsp;</p>
<p align="center" class="style2"><img src="logo.jpg" width="637" height="83"></p>
<p align="center" class="style3">&nbsp;</p>
<form name="form1" method="post" action="gera_consulta_curso.asp">
  <table width="53%" border="0" align="center">
    <tr> 
      <td width="47%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099">Selecione 
        o Mega-Processo</font></td>
      <td width="53%"> 
        <select name="selMega" id="selMega">
          <option value="0" selected>== TODOS ==</option>
          <%
		do until rs.eof=true
	%>
          <option value="<%=trim(rs("MEPR_TX_ABREVIA_CURSO"))%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
		rs.movenext
	loop
	%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="47%" height="15"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099">ou 
        Frente</font></td>
      <td width="53%" height="15"><b><font color="#FF0000"><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif">
        <select name="selFrente">
          <option value="0" selected>== TODOS ==</option>
          <option value="1">BW</option>
          <option value="2">FI</option>
          <option value="3">OIL & CO</option>
          <option value="4">P* & MES</option>
          <option value="5">RH</option>
        </select>
        </font></font></font></b></td>
    </tr>
    <tr> 
      <td width="47%" height="15"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099"></font></td>
      <td width="53%" height="15">&nbsp;</td>
    </tr>
    <tr> 
      <td width="47%" height="18"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099">Selecione 
        a Abrang&ecirc;ncia</font></td>
      <td width="53%" height="18"> 
        <select name="selAbrangencia">
          <option value="6,8,9" selected>== TODOS ==</option>
          <option value="6,9">PETROBRAS</option>
          <option value="8,9">REFAP</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="47%" height="12">&nbsp;</td>
      <td width="53%" height="12">&nbsp;</td>
    </tr>
    <tr> 
      <td width="47%" height="22"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099">Selecione 
        a Situa&ccedil;&atilde;o </font></td>
      <td width="53%" height="22"> <b> <font color="#FF0000"> <font size="1"> 
        <font face="Verdana, Arial, Helvetica, sans-serif"> 
        <select name="selStatus">
          <option value="0" selected>== TODOS ==</option>
          <option value="1">CURSOS EM ATRASO</option>
          <option value="2">EM APROVAÇÃO TREINAMENTO</option>
          <option value="3">PUBLICADO EM PRODUÇÃO</option>
        </select>
        </font></font></font></b></td>
    </tr>
    <tr> 
      <td width="47%" height="13"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099"></font></td>
      <td width="53%" height="13">&nbsp; </td>
    </tr>
  </table>
  <div align="center"></div>
  <p align="center">
    <input type="submit" name="Submit" value="Visualizar" class="boton_box">
  </p>
  <p align="center"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000066"><b><font color="#000099">Criado 
    por Gest&atilde;o do Conhecimento</font></b></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>
    <font color="#000099">Total de Visitas : <b><%=visitas%></b></font></font></p>
</form>
<p>&nbsp;</p>
</body>

<%
rs.close
set rs = nothing

db.close
set db = nothing
%>	
	
</html>
