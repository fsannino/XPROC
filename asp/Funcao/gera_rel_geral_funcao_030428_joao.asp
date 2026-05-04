<%@LANGUAGE="VBSCRIPT"%> 
 
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Funcao
dim str_Modulo

str_DsModulo = request("txtDescModulo")

str_MegaProcesso = request("selMegaProcesso")
str_Modulo=request("selSubModulo")
'response.Write(str_Modulo)
'if str_modulo<>0 and str_modulo<>"" then
'	compl1=" AND SUMO_NR_SEQUENCIA=" & str_modulo 
'end if

if str_modulo<>"" then
	compl1=" AND SUMO_NR_SEQUENCIA=" & str_modulo 
end if

str_Opc = Request("txtOpc")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso)

if str_MegaProcesso=0 then
	ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO ORDER BY FUNE_TX_TITULO_FUNCAO_NEGOCIO"
else
	ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & compl1 & " ORDER BY FUNE_TX_TITULO_FUNCAO_NEGOCIO"
end if

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
//  End -->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
Geral de Fun&ccedil;&atilde;o R/3</b></font></p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </font></p>
 <p style="margin-top: 0; margin-bottom: 0"><font size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000">
 (<font face="Verdana">Clique no código da Fun&ccedil;&atilde;o R/3 para exibir
 seus dados)</font></font></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <%
        conta=0
        DO UNTIL RS.EOF=TRUE%>
        <table border="0" width="74%" height="62">
          <tr>
            <td width="12%" height="19">
            </td>
            <td width="16%" height="19" bgcolor="#E0E0E0">
              <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#330099"><b>Código</b></font></td>
            <td width="90%" height="19">
              <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="1"><a href="exibe_dados_funcao.asp?selMegaProcesso=<%=str_MegaProcesso%>&selFuncao=<%=RS("FUNE_CD_FUNCAO_NEGOCIO")%>"><b><%=RS("FUNE_CD_FUNCAO_NEGOCIO")%></b></a></font></td>
          </tr>
          <tr>
            <td width="12%" height="19"></td>
            <td width="16%" height="19" bgcolor="#E0E0E0"><font face="Verdana" size="1" color="#330099"><b>Título</b></font></td>
            <td width="90%" height="19"><font face="Verdana" color="#330099" size="1"><%=RS("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
          </tr>
          <tr>
            <td width="12%" height="6" valign="top"></td>
            <td width="16%" height="6" valign="top" bgcolor="#E0E0E0"><font face="Verdana" size="1" color="#330099"><b>Descrição</b></font></td>
            <td width="90%" height="6" valign="top"><font face="Verdana" color="#330099" size="1"><%=RS("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></font></td>
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
 <font face="Verdana" size="2" color="#800000">Não existe Fun&ccedil;&atilde;o R/3s para o Mega-Processo Selecionado</font>
 <%end if%>
 </b>
  <table width="75%" border="0">
    <tr>
      <td width="16%">&nbsp;</td>
      <td width="84%"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
        de Fun&ccedil;&otilde;es Listadas</strong> : <%=conta%></font></td>
    </tr>
  </table>
  </form>
</body>
</html>
