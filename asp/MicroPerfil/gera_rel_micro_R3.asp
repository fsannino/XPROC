<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Mega = request("selMegaProcesso")
str_Macro = request("selMacroPerfil")
COMPL = ""

'response.Write("<p>Mega=" & str_Mega)
'response.Write("<p>Macro=" & str_Macro)

if str_Mega <> 0 then
	COMPL = COMPL & " and MACRO_PERFIL.MEPR_CD_MEGA_PROCESSO =" & str_Mega
else
	COMPL = COMPL & ""
end if
if str_Macro <> 0 then
	COMPL = COMPL & " and dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL =" & str_Macro
else
	COMPL = COMPL & ""
end if

'str_SQl = "SELECT * FROM " & Session("PREFIXO") & "MICRO_PERFIL_R3" & COMPL &" ORDER BY MEPR_CD_MEGA_PROCESSO, MIPE_NR_SEQ_MICRO_PERFIL"

str_SQL = "" 
str_SQL = str_SQL & " SELECT "
str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3.MIPE_NR_SEQ_MICRO_PERFIL, dbo.MICRO_PERFIL_R3.MIPE_TX_NOME_TECNICO, "
str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3.MIPE_TX_DESC_MICRO_PERFIL, dbo.MICRO_PERFIL_R3.MIPE_TX_DESC_DETALHADA, "
str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL, dbo.MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, "
str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, dbo.MACRO_PERFIL.MCPE_TX_NOME_TECNICO "
str_SQL = str_SQL & " FROM dbo.MICRO_PERFIL_R3 INNER JOIN"
str_SQL = str_SQL & " dbo.MACRO_PERFIL ON "
str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL INNER JOIN"
str_SQL = str_SQL & " dbo.MEGA_PROCESSO ON dbo.MACRO_PERFIL.MEPR_CD_MEGA_PROCESSO = dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO INNER JOIN"
str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO ON dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO"
str_SQL = str_SQL & " where MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO > 0 "
'response.Write(str_SQL & COMPL)
set rs=conn_db.execute(str_SQl & COMPL)
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

</head>
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


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="gera_rel_micro.asp">
  <input type="hidden" name="INC" size="20" value="1"> 
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr>
            <td bgcolor="#330099" width="39" valign="middle" align="center">
              <div align="center">
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Cenario/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Cenario/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Cenario/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../Cenario/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Cenario/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../Cenario/home.gif" border="0"></a>&nbsp;</div>
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
  <table width="87%" border="0" cellpadding="0" cellspacing="5" name="tblSubProcesso" height="47">
    <tr> 
      <td width="7%" height="1"></td>
      <td width="83%" height="1"> 
      </td>
      <td width="5%" height="1"> 
      </td>
      <td width="17%" height="1"> 
      </td>
    </tr>
    <tr> 
      <td width="7%" height="1">&nbsp;</td>
      <td width="83%" height="1"> 
        <input type="hidden" name="txtOpc" value="1">
        <p align="left"><font color="#330099" face="Verdana" size="3">Relatório 
          de Micro-Perfil - Criado no R/3</font></td>
      <td width="5%" height="1"> 
       </td>
      <td width="17%" height="1"> 
       </td>
    </tr>
  </table>
  &nbsp;
  <%
  tem=0
  do until rs.eof=true
  tem=1
  %>
  <table border="0" width="61%" cellpadding="2">
    <tr> 
      <td width="27%" bgcolor="#330099"><font face="Verdana" size="1" color="#FFFFFF"><b>Mega-Processo</b></font></td>
      <td width="73%"><font face="Verdana" size="1"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#330099"><font face="Verdana" size="1" color="#FFFFFF"><b>Macro-Perfil</b></font></td>
      <td><font face="Verdana" size="1"><strong><%=rs("MCPE_TX_NOME_TECNICO")%> </strong></font></td>
    </tr>
    <tr> 
      <td bgcolor="#330099"><font face="Verdana" size="1" color="#FFFFFF"><b>Função 
        R/3</b></font></td>
      <td><font face="Verdana" size="1"><%=rs("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
    </tr>
    <tr> 
      <td width="27%" bgcolor="#330099"><font face="Verdana" size="1" color="#FFFFFF"><b>Micro-Perfil</b></font></td>
      <td width="73%"><font face="Verdana" size="1"><strong><%=rs("MIPE_TX_NOME_TECNICO")%></strong></font></td>
    </tr>
    <tr> 
      <td width="27%" bgcolor="#330099"><font face="Verdana" size="1" color="#FFFFFF"><b>Descrição</b></font></td>
      <td width="73%"><font face="Verdana" size="1"><%=rs("MIPE_TX_DESC_MICRO_PERFIL")%></font></td>
    </tr>
    <tr> 
      <td width="27%" bgcolor="#330099"><font face="Verdana" size="1" color="#FFFFFF"><b>Descrição 
        Detalhada </b></font></td>
      <td width="73%"><font face="Verdana" size="1"><%=rs("MIPE_TX_DESC_DETALHADA")%></font></td>
    </tr>
  </table>
  <p>
  <%
  rs.movenext
  loop
  if tem=0 then
  %>
  <p><font color="#800000"><b>Nenhum Registro Encontrado para a Seleção!</b></font></p>
  <%end if%>
  </form>

</body>
</html>
