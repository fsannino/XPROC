<<<<<<< HEAD
 
<%
str_CdUsuario=request("ID")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql="SELECT * FROM " & Session("PREFIXO") & "USUARIO WHERE USUA_CD_USUARIO='" & trim(str_CdUsuario) & "'"
set rs2=db.execute(ssql)

VALOR=RS2("USUA_TX_NOME_USUARIO")

ssql="SELECT * FROM " & Session("PREFIXO") & "ACESSO WHERE USUA_CD_USUARIO ='" & trim(str_CdUsuario) & "'"

set rs=db.execute(ssql)

%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">

function fechar()
{
window.close();
}

function mover()
{
window.moveTo(100	,200)
}

</script>

<body topmargin="0" leftmargin="0" onload="javascript:mover()">
<table border="0" width="69%">
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="71%"><font face="Verdana" size="2" color="#000080"><b>Relação
      de </b></font><font face="Verdana" size="2" color="#000080"><b>Mega-Processos
      associados</b></font></td>
    <td width="5%">
    </td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"></td>
  </tr>
  <tr>
    <td width="24%"><font face="Verdana" size="2" color="#000080"><b>Usu&aacute;rio 
      :</b></font></td>
    <td width="76%" colspan="2"><font face="Verdana" size="2" color="#000080"><%=valor%></font></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"><font face="Verdana" size="2">
    <select size="8" name="D1" multiple>
      <%if rs.eof=true then%>
      <option>Nenhum Mega-Processo associado a esta classe</option>
	   <%else%>
		<%DO UNTIL RS.EOF=TRUE
	   IF NOT ISNULL(rs("MEPR_CD_MEGA_PROCESSO")) THEN
	   SSQL1="SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO")
	   SET RS1=DB.EXECUTE(SSQL1)
	   VALOR2=RS1("MEPR_TX_DESC_MEGA_PROCESSO")
	   END IF
	   %>
      <option><%=VALOR2%></option>
      <%
      RS.MOVENEXT
      LOOP
      end if
      %>
   </select></font></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2">
      <p align="right"><b><a href="#" onclick="javascript:fechar()"><font face="Verdana" size="2" color="#800000">Fechar
      Janela</font></a></b></td>
  </tr>
</table>

</body>

</html>
=======
 
<%
str_CdUsuario=request("ID")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql="SELECT * FROM " & Session("PREFIXO") & "USUARIO WHERE USUA_CD_USUARIO='" & trim(str_CdUsuario) & "'"
set rs2=db.execute(ssql)

VALOR=RS2("USUA_TX_NOME_USUARIO")

ssql="SELECT * FROM " & Session("PREFIXO") & "ACESSO WHERE USUA_CD_USUARIO ='" & trim(str_CdUsuario) & "'"

set rs=db.execute(ssql)

%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">

function fechar()
{
window.close();
}

function mover()
{
window.moveTo(100	,200)
}

</script>

<body topmargin="0" leftmargin="0" onload="javascript:mover()">
<table border="0" width="69%">
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="71%"><font face="Verdana" size="2" color="#000080"><b>Relação
      de </b></font><font face="Verdana" size="2" color="#000080"><b>Mega-Processos
      associados</b></font></td>
    <td width="5%">
    </td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"></td>
  </tr>
  <tr>
    <td width="24%"><font face="Verdana" size="2" color="#000080"><b>Usu&aacute;rio 
      :</b></font></td>
    <td width="76%" colspan="2"><font face="Verdana" size="2" color="#000080"><%=valor%></font></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"><font face="Verdana" size="2">
    <select size="8" name="D1" multiple>
      <%if rs.eof=true then%>
      <option>Nenhum Mega-Processo associado a esta classe</option>
	   <%else%>
		<%DO UNTIL RS.EOF=TRUE
	   IF NOT ISNULL(rs("MEPR_CD_MEGA_PROCESSO")) THEN
	   SSQL1="SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO")
	   SET RS1=DB.EXECUTE(SSQL1)
	   VALOR2=RS1("MEPR_TX_DESC_MEGA_PROCESSO")
	   END IF
	   %>
      <option><%=VALOR2%></option>
      <%
      RS.MOVENEXT
      LOOP
      end if
      %>
   </select></font></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2"></td>
  </tr>
  <tr>
    <td width="24%"></td>
    <td width="76%" colspan="2">
      <p align="right"><b><a href="#" onclick="javascript:fechar()"><font face="Verdana" size="2" color="#800000">Fechar
      Janela</font></a></b></td>
  </tr>
</table>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
