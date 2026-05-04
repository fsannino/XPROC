<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn_consulta.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3
%>

<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Concessão de Perfil de Acesso</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function manda()
{
if(document.frm1.txtchave.value=="")
{
alert('Você deve digitar uma chave!');
document.frm1.txtchave.focus();
return;
}
else
{
document.frm1.submit();
}
}
</script>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<script language="javascript" src="troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<form name="frm1" method="POST" action="gera_consulta_gestor.asp" target="_top">
   <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; </p>
   <table border="0" width="75%">
              <tr><td width="71%">
<p align="center"><b><font face="Verdana" color="#000080">Verificação de Gestor de Pessoas</font></b></p>
                         </td>
              </tr>
		   </table>
&nbsp;<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="74%" id="AutoNumber1">
              <tr>
                         <td width="18%">&nbsp;</td>
                         <td width="22%">&nbsp;</td>
                         <td width="60%">&nbsp;</td>
              </tr>
              <tr>
                         <td width="18%">&nbsp;</td>
                         <td width="22%"><b><font face="Verdana" color="#000080" size="2">Digite a Chave</font></b></td>
                         <td width="60%"><input type="text" name="txtchave" size="12" maxlength="4"></td>
              </tr>
              <tr>
                         <td width="18%">&nbsp;</td>
                         <td width="22%">&nbsp;</td>
                         <td width="60%">&nbsp;</td>
              </tr>
              <tr>
                         <td width="18%">&nbsp;</td>
                         <td width="22%">&nbsp;</td>
                         <td width="60%">
                         <input type="button" value="Consultar" name="B1" onClick="manda()">
                         </td>
              </tr>
   </table>
   <p>&nbsp;</p>
</form>
</body>

</html>

<script>
document.frm1.txtchave.focus();
</script>