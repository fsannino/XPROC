<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

trans = request("transacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO_MEGA WHERE TRAN_CD_TRANSACAO='" & TRANS & "'")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Exibição de Dono</title>
</head>
<script>
function move()
{
window.moveTo(50,50)
}
</script>
<body link="#800000" vlink="#800000" alink="#800000" onload="javascript:move()">

<table border="0" width="306">
  <tr>
    <td width="296">

<p><font size="2" face="Verdana" color="#000080"><b>Transação Selecionada</b>
: <%=request("transacao")%></font></p>
<p><select size="4" name="donos" multiple>
<%if rs.eof=true then%>
<option>== Transação sem Dono Definido ==</option>
<%
else
do until rs.eof=true
SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
%>
<option><%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
<%
rs.movenext
loop
END IF
%>
&nbsp;
</select></p>

<p align="left"><font color="#800000"><b><a href="javascript:window.close()">Fechar
Janela</a></b></font></p>

    </td>
  </tr>
</table>

</body>

</html>
=======
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

trans = request("transacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO_MEGA WHERE TRAN_CD_TRANSACAO='" & TRANS & "'")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Exibição de Dono</title>
</head>
<script>
function move()
{
window.moveTo(50,50)
}
</script>
<body link="#800000" vlink="#800000" alink="#800000" onload="javascript:move()">

<table border="0" width="306">
  <tr>
    <td width="296">

<p><font size="2" face="Verdana" color="#000080"><b>Transação Selecionada</b>
: <%=request("transacao")%></font></p>
<p><select size="4" name="donos" multiple>
<%if rs.eof=true then%>
<option>== Transação sem Dono Definido ==</option>
<%
else
do until rs.eof=true
SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
%>
<option><%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
<%
rs.movenext
loop
END IF
%>
&nbsp;
</select></p>

<p align="left"><font color="#800000"><b><a href="javascript:window.close()">Fechar
Janela</a></b></font></p>

    </td>
  </tr>
</table>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
