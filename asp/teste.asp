<<<<<<< HEAD
 
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Nova pagina 1</title>
</head>

<body>
<table border="1" width="26%">
<%DO UNTIL RS.EOF=TRUE%>
  <tr>
    <td width="49%">&nbsp;</td>
    <td width="51%">xxx</td>
    <td width="51%"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%>xxx</td>
  </tr>
<%
RS.MOVENEXT
LOOP
%>
  <tr>
    <td width="49%">yyy</td>
    <td width="51%"><input type="checkbox" name="C1" value="ON"></td>
    <td width="51%"><input type="checkbox" name="C1" value="ON"></td>
  </tr>
</table>
</body>

</html>
=======
 
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Nova pagina 1</title>
</head>

<body>
<table border="1" width="26%">
<%DO UNTIL RS.EOF=TRUE%>
  <tr>
    <td width="49%">&nbsp;</td>
    <td width="51%">xxx</td>
    <td width="51%"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%>xxx</td>
  </tr>
<%
RS.MOVENEXT
LOOP
%>
  <tr>
    <td width="49%">yyy</td>
    <td width="51%"><input type="checkbox" name="C1" value="ON"></td>
    <td width="51%"><input type="checkbox" name="C1" value="ON"></td>
  </tr>
</table>
</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
