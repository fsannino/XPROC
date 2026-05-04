<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%

'Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"

'if Right(Session("CdUsuario"),2) = "FI" or Right(Session("CdUsuario"),2) = "CO" OR Session("CdUsuario") = "G1FI" or Session("CdUsuario") = "G1CO" or Session("CdUsuario") = "XXXX" or Session("CdUsuario") = "XK45" or Session("CdUsuario") = "XD47" or Session("CdUsuario") = "X939" or str_Chave = "XK45" or str_Chave = "XD47" or str_Chave = "X939" then
   Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest001;uid=cogest;database=cogest"
   Session("Conn_String_Cronograma_Gravacao")= "Provider=SQLOLEDB.1;server=10.22.22.13;pwd=sinergia;uid=sinergia;database=sinergia"

'else
'   'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"
'   Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"
'end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/biblioteca.css" rel="stylesheet" type="text/css">
<link href="css/objinterface.css" rel="stylesheet" type="text/css">
</head>

<!--#include file="includes/include_TopoMenu.asp" -->
<table width="75%" border="0" cellspacing="10">
  <tr> 
    <td width="29%"><div align="right"></div></td>
    <td width="61%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="right"></div></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="right">Onda:</div></td>
    <td>
      <!--#include file="asp/includes_old/inc_Combo_Onda.asp" -->
    </td>
    <td>&nbsp; </td>
  </tr>
  <tr> 
    <td><div align="right">Plano:</div></td>
    <td>&nbsp; </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="right">Atividades:</div></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="75%" border="0">
<!--#include file="includes/includerodape.asp" -->
</table>
</body>
</html>
