 
<%
if request("popt") <> "" then
   str_Opt = request("popt")
else
   str_Opt = ""
end if

if request("selMegaProcesso") <> "" then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = 0
end if

'response.write str_MegaProcesso

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQL = str_SQL & " SELECT MCPR_NR_SEQ_MACRO_PERFIL, "
str_SQL = str_SQL & " MCPE_TX_NOME_TECNICO, "
str_SQL = str_SQL & " MCPE_TX_DESC_MACRO_PERFIL"
str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_PERFIL"
str_SQL = str_SQL & " WHERE ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"  
'str_SQL = str_SQL & " AND MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL = str_SQL & " AND MCPE_TX_SITUACAO = 'NA'"
'response.write str_SQL
set rs_Macro=db.execute(str_SQL)

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

function Confirma()
{
   if(document.frm1.selMegaProcesso.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de um Mega Processo !");
      document.frm1.selMegaProcesso.focus();
      return;
      }
   else
      {
      document.frm1.submit();
      }		
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="POST" action="grava_macro_perfil.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top">
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a>
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <div align="center"><font face="Verdana" color="#330099" size="3">Rela&ccedil;&atilde;o 
          de Macros n&atilde;o aprovados</font></div>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="13%">&nbsp;</td>
      <td width="23%" bgcolor="#330099"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Nome 
        T&eacute;cnico</font></b></td>
      <td width="52%" bgcolor="#330099"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Descri&ccedil;&atilde;o</font></b></td>
      <td width="8%">&nbsp;</td>
    </tr>
    <% if not rs_Macro.EOF then 
	      int_sequencia = 0
		  str_Cor1 = "#FFFFFF"
	      Do while not rs_Macro.EOF 
		     int_sequencia = int_sequencia + 1		  
	  IF str_Cor1 = "#FFFFFF" then 
	        str_Cor1 = "#EFEFEF"
			str_Cor2 = "#FFFFFF"
		 else
	        str_Cor1 = "#FFFFFF"
			str_Cor2 = "#EFEFEF" 
		 end if 	

	%>
    <tr> 
      <td width="13%">&nbsp;</td>
      <td width="23%" bgcolor="<%=str_Cor2%>"> 
        <div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"> 
          <input type="hidden" name="txtMacro<%=int_sequencia%>" value="<%=rs_Macro("MCPR_NR_SEQ_MACRO_PERFIL")%>">
          &nbsp;&nbsp;&nbsp; <%=rs_Macro("MCPE_TX_NOME_TECNICO")%></font></div>
      </td>
      <td width="52%" height="30" bgcolor="<%=str_Cor2%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%=rs_Macro("MCPE_TX_DESC_MACRO_PERFIL")%></font></td>
      <td width="8%">&nbsp;</td>
    </tr>
    <% rs_Macro.Movenext
	Loop %>
    <tr> 
      <td width="13%">&nbsp;</td>
      <td width="23%">&nbsp;</td>
      <td width="52%">&nbsp;</td>
      <td width="8%">&nbsp;</td>
    </tr>
    <% else %>
    <tr> 
      <td width="13%">&nbsp;</td>
      <td width="23%">&nbsp;</td>
      <td width="52%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">N&atilde;o 
        encontrado Macro Perfil com situa&ccedil;&atilde;o de n&atilde;o aprovado.</font></td>
      <td width="8%">&nbsp;</td>
    </tr>
    <% end if %>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
