<%@LANGUAGE="VBSCRIPT"%> 
<%

strMSG = Request("pMsg")
'response.Write(strMSG)
if strMSG = "C" then
   str_Tipo = "C"
   strMSG = "Acesso apenas para consulta"
end if

strPlano =  Request("pPlano")
strUsuario = Request("pUsua")

strErroServidor = Request("pErroServidor")

set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
'conn_Cogest.CursorLocation=3

if request("pLote") <> 0 then
   str_Lote = request("pLote")
else
   str_Lote = 0
end if

'conn_Cogest.BeginTransaction
str_SQL = ""
str_SQL = str_SQL & " DELETE FROM dbo.GOLI_FUNCAO_USUARIO_SEM_PERFIL "
str_SQL = str_SQL & " WHERE LOTE_NR_SEQ_LOTE = " & str_Lote
conn_Cogest.Execute(str_SQL)

str_SQL = ""
str_SQL = str_SQL & " DELETE FROM dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL "
str_SQL = str_SQL & " WHERE LOTE_NR_SEQ_LOTE = " & str_Lote
conn_Cogest.Execute(str_SQL)

str_SQL = ""
str_SQL = str_SQL & " DELETE FROM dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL_RH "
str_SQL = str_SQL & " WHERE LOTE_NR_SEQ_LOTE = " & str_Lote
conn_Cogest.Execute(str_SQL)

str_SQL = ""
str_SQL = str_SQL & " DELETE FROM dbo.GOLI_LOTE "
str_SQL = str_SQL & " WHERE LOTE_NR_SEQ_LOTE = " & str_Lote
conn_Cogest.Execute(str_SQL)
'conn_Cogest.Commit
%>
<html>
<head>
<script>
function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href.href='altera_Atividade1.asp?selAtiv='+document.frm1.selAtividade.value
	 }
 }
</SCRIPT>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top">&nbsp;      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20">&nbsp; </td>
    </tr>
  </table>
  <table width="627" height="150" border="0" cellpadding="5" cellspacing="5">
    <tr>
      <td height="29"></td>
      <td height="29" valign="middle" align="left"></td>
      <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td height="29"></td>
      <td height="29" valign="middle" align="left"></td>
      <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td width="101" height="29"></td>
      <td width="125" height="29" valign="middle" align="left"></td>
      <td height="29" valign="middle" align="left" colspan="2">&nbsp;        </td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2">&nbsp; </td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="51"><a href="../../indexA.asp"><img src="../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
      <td height="1" valign="middle" align="left" width="458"><font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
    </tr>
    <tr>
      <td width="101" height="1"></td>
      <td width="125" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
  </table>
</form>
</body>
</html>
