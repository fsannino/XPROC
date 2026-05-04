<%@LANGUAGE="VBSCRIPT"%> 
 
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Funcao
dim str_Modulo

str_MegaProcesso = request("selMegaProcesso")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso)

ssql="SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " ORDER BY MCPE_TX_NOME_TECNICO"

'response.write ssql

set rs=db.execute(ssql)
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}

function ver_historico(macro)
{
var a=macro;
window.open("ver_historico.asp?macro=" + a + "","_blank","width=600,height=260,history=0,scrollbars=1,titlebar=0,resizable=0,top=150,left=300")
}

//  End -->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF">
<form name="frm1" method="post" action="">
 <input type="hidden" name="txtOpc" value="1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Relatório
Geral de Macro - Perfil</b></font></p>
<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Mega-Processo
Selecionado : <%=str_MegaProcesso%> - <%=rs_mega("MEPR_TX_DESC_MEGA_PROCESSO")%></font></p>
 <p style="margin-top: 0; margin-bottom: 0"><font size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000">
 (<font face="Verdana">Clique no código do Macro-Perfil para exibir as
 transações relacionadas)</font></font></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <%
        conta=0
        DO UNTIL RS.EOF=TRUE%>
        
  <table border="0" width="74%" height="50">
    <tr> 
      <td width="12%" height="19"> </td>
      <td width="23%" height="19" bgcolor="#E0E0E0"> <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>Nome 
          Técnico</b></font></td>
      <td width="83%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="1"><a href="exibe_transacao_macro.asp?selMegaProcesso=<%=str_MegaProcesso%>&txtOPT=3&selMacroPerfil=<%=RS("MCPR_NR_SEQ_MACRO_PERFIL")%>"><b><font size="2"><%=RS("MCPE_TX_NOME_TECNICO")%></font></b></a> <a href="javascript:ver_historico('<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>')"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico"></a><a href="#" onclick="ver_historico('<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>')"></a></font></td>
    </tr>
    <tr> 
      <td width="12%" height="1" valign="top"></td>
      <td width="23%" height="1" valign="top" bgcolor="#E0E0E0"><font face="Verdana" size="2" color="#330099"><b>Descrição</b></font></td>
      <td width="83%" height="1" valign="top"><font face="Verdana" color="#330099" size="2"><%=RS("MCPE_TX_DESC_MACRO_PERFIL")%> </font></td>
    </tr>
    <tr>
      <td height="1" valign="top"></td>
	    <% str_Situacao = RS("MCPE_TX_SITUACAO")
		If str_Situacao = "EE" then
			str_Situacao = "Em elaboração"
		 elseIf str_Situacao = "AT" then
			str_Situacao = "Alterado transação"
		 elseIf str_Situacao = "EA" then
			str_Situacao = "Em aprovação"			  
		 elseIf str_Situacao = "NA" then
			str_Situacao = "Não aprovado"			  
		 elseIf str_Situacao = "EC" then
			str_Situacao = "Em criação no R/3"			  
		 elseIf str_Situacao = "RE" then
			str_Situacao = "Recusado no R/3"			  
		 elseIf str_Situacao = "EX" then
			str_Situacao = "Excluída a função"			  
		 elseIf str_Situacao = "MR" then
			str_Situacao = "Mudado para referência"			  
		 elseIf str_Situacao = "EL" then
			str_Situacao = "Excluído"			  
		 elseIf str_Situacao = "CR" then
			str_Situacao = "Criado no R3"			  
		 elseIf str_Situacao = "AR" then
			str_Situacao = "Em alteração no R/3"			  
		 elseIf str_Situacao = "ER" then
			str_Situacao = "Em exclusão no R/3"			  
		 elseIf str_Situacao = "AP" then
			str_Situacao = "Alterado no R/3"			  
		 elseIf str_Situacao = "EP" then
			str_Situacao = "Excluído no R/3"			  
         end if
	  %>
      <td height="1" valign="top" bgcolor="#E0E0E0"><font face="Verdana" size="2" color="#330099"><b>Status</b></font></td>
      <td height="1" valign="top"><font face="Verdana" color="#330099" size="2"><%=str_Situacao%></font></td>
    </tr>
  </table>
 <P>
        <%
        conta=conta+1
			RS.MOVENEXT        
        	LOOP
        %>
 <b>
 <%if conta=0 then%>
 &nbsp;<font face="Verdana" size="2" color="#800000">&nbsp;&nbsp;&nbsp;</font>&nbsp;<font face="Verdana" size="2" color="#800000">&nbsp;&nbsp;&nbsp;</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 <font face="Verdana" size="2" color="#800000">Não existe Macro-Perfil para o Mega-Processo Selecionado</font>
 <%end if%>
 </b>
<p>&nbsp;</p>
  </form>
</body>
</html>
