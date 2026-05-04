<%@LANGUAGE="VBSCRIPT"%>  
<%

if request("opt") = 1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
elseif request("opt") = 2 then
	Response.Buffer = TRUE
	Response.ContentType = "application/msword"
end if

str_Opc = Request("txtOpc")

str_Funcao=request("selFuncao")

str_MegaProcesso = "0"
if request("txtMegaProcesso") <> "" then
   str_MegaProcesso = request("txtMegaProcesso")
else
   if request("selMegaProcesso") <> "" then
      str_MegaProcesso = request("selMegaProcesso")      
   end if	  
end if   

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Sql = ""
str_Sql = str_Sql & " SELECT "
str_Sql = str_Sql & " dbo.MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_TX_DESC_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " FROM dbo.FUNCAO_NEGOCIO INNER JOIN"
str_Sql = str_Sql & " dbo.MEGA_PROCESSO ON dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_Sql & " where dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'"  
set rds_Funcao=db.execute(str_Sql)

str_Sql = ""
str_Sql = str_Sql & " SELECT dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " , dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
str_Sql = str_Sql & " , dbo.MICRO_PERFIL_R3.MIPE_NR_SEQ_MICRO_PERFIL "
str_Sql = str_Sql & " , dbo.MICRO_PERFIL_R3.MIPE_TX_NOME_TECNICO"
str_Sql = str_Sql & " , dbo.MICRO_PERFIL_R3.MIPE_TX_DESC_MICRO_PERFIL"
str_Sql = str_Sql & " , dbo.MICRO_PERFIL_R3.MIPE_TX_DESC_DETALHADA"
str_Sql = str_Sql & " , dbo.MICRO_PERFIL_R3.MIPE_TX_ORIENTACAO"
str_Sql = str_Sql & " FROM dbo.FUNCAO_NEGOCIO INNER JOIN"
str_Sql = str_Sql & " dbo.MACRO_PERFIL ON dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO INNER JOIN"
str_Sql = str_Sql & " dbo.MICRO_PERFIL_R3 ON dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL"
str_Sql = str_Sql & " where dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'"  
set rds_Perfil=db.execute(str_Sql)
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
function Confirma()
{
   document.frm1.submit(); 
}

//  End -->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="valida_ori_massa_perfil.asp">
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
      <td colspan="3" height="20"> <table width="625" border="0" align="center">
          <tr> 
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
            <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
            <td width="195"></td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28"></td>
            <td width="26">&nbsp;</td>
            <td width="159"></td>
          </tr>
        </table></td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    Cadastro de Observa&ccedil;&atilde;o Espec&iacute;fica para Perfil 
    R/3</b></font></p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    </font></p>
 <p style="margin-top: 0; margin-bottom: 0"><font size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000">
 (<font face="Verdana">Clique no código da Fun&ccedil;&atilde;o R/3 para exibir
 seus dados)</font></font></font></p>
  <table width="87%" border="0">
    <tr>
      <td width="3%">&nbsp;</td>
      <td width="77%">&nbsp;</td>
      <td width="20%"><a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>"><img src="../../imagens/conteudo_01.gif" width="18" height="22" border="0"></a><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>">Relat&oacute;rio 
        completo</a> </font></td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>        
  <table width="87%" height="86" border="0" cellpadding="3" cellspacing="1">   
    <tr valign="top"> 
      <td width="12%"> <p align="right" style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="2">Mega:</font></strong></td>
      <td width="88%"> <p style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="1"><%=rds_Funcao("MEPR_TX_DESC_MEGA_PROCESSO")%> </font></strong></td>
    </tr>	
    <tr valign="top"> 
      <td><div align="right"><strong><font face="Verdana" size="2">Fun&ccedil;&atilde;o:</font></strong></div></td>
      <td><strong><font face="Verdana" size="1"><%=rds_Funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rds_Funcao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%> 
        </font></strong></td>
    </tr>
    <tr valign="top"> 
      <td><div align="right"><font face="Verdana" size="1">&nbsp;<strong><font face="Verdana" size="2">Descrição:</font></strong></font></div></td>
      <td><font face="Verdana" size="1"><%=rds_Funcao("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></font></td>
    </tr>
    <tr valign="top" bgcolor="#0000FF"> 
      <td height="5" bgcolor="#000099"></td>
      <td height="5" bgcolor="#000099"></td>
    </tr>
  </table>
          <table width="87%" height="129" border="1" cellpadding="1" cellspacing="1">
       <%
  conta=0
  int_Contador = 0 
  Do While not rds_Perfil.EOF
     int_Contador = int_Contador + 1
		%>
            <tr valign="top">
              <td height="19"><strong><font face="Verdana" size="2">Perfil</font></strong></td>
              <td height="19"><strong><font face="Verdana" size="1"><%=rds_Perfil("MIPE_TX_DESC_MICRO_PERFIL")%>
                      <input name="txtSeqMacro<%=int_Contador%>" type="hidden" id="txtSeqMacro<%=int_Contador%>" value="<%=rds_Perfil("MCPR_NR_SEQ_MACRO_PERFIL")%>">
                      <input name="txtSeqPerfil<%=int_Contador%>" type="hidden" id="txtSeqPerfil<%=int_Contador%>" value="<%=rds_Perfil("MIPE_NR_SEQ_MICRO_PERFIL")%>">
</font></strong></td>
            </tr>
            <tr valign="top">
              <td width="17%" height="19"><font face="Verdana" size="1">&nbsp;<strong><font face="Verdana" size="2">Descri&ccedil;&atilde;o</font></strong></font></td>
              <td width="83%" height="19"><font face="Verdana" size="1"><%=rds_Perfil("MIPE_TX_DESC_DETALHADA")%></font></td>
            </tr>
            <tr valign="top">
              <td height="19"><strong><font face="Verdana" size="2">Observa&ccedil;&atilde;o Espec&iacute;fica para Perfil </font></strong></td>
              <td height="19"><font face="Verdana" size="1"><strong><font face="Verdana" size="1">
                <textarea name="txtDescPerfil<%=int_Contador%>" cols="90" rows="5"><%=rds_Perfil("MIPE_TX_ORIENTACAO")%></textarea>
                <input name="txtObsAnterior<%=int_Contador%>" type="hidden" id="txtObsAnterior<%=int_Contador%>" value="<%=rds_Perfil("MIPE_TX_ORIENTACAO")%>">
              </font></strong></font></td>
            </tr>
            <tr valign="top" bgcolor="#0000FF">
              <td height="5" bgcolor="#000099"></td>
              <td height="5" bgcolor="#000099"></td>
            </tr>
            <%
        conta=conta+1
		rds_Perfil.movenext        
   	Loop
        %>
          </table>
          <input name="txtQtdObj" type="hidden" id="txtQtdObj" value="<%=int_Contador%>">
 <P>
 <b>
 <%if conta=0 then%>
 &nbsp;<font face="Verdana" size="2" color="#800000">&nbsp;&nbsp;&nbsp;</font>&nbsp;<font face="Verdana" size="2" color="#800000">&nbsp;&nbsp;&nbsp;</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 <font face="Verdana" size="2" color="#800000">Não existe Perfil para a Fun&ccedil;&atilde;o Selecionada</font>
 <%end if%>
 </b>
  <table width="75%" border="0">
    <tr>
      <td width="16%">&nbsp;</td>
      <td width="84%"><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Total 
        de Perfil Listado</strong> : <%=conta%></font></td>
    </tr>
  </table>
</form>
</body>
</html>
