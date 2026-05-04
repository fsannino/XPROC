<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<SCRIPT LANGUAGE="JavaScript">

function Confirma()
{
document.frm1.submit();
}

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

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#330099" vlink="#330099" alink="#330099">
<form name="frm1" method="post" action="valida_status3.asp">
        <font color="#330099" face="Verdana" size="2"> <input type="hidden" name="mega" size="20" value="<%=request("selMegaProcesso")%>"></font><input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
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
      &nbsp;
    </td>
  </tr>
</table>
        <p align="center"><font color="#330099" face="Verdana" size="3"><b>Encaminhamento
        de Status de Macro-Perfil</b></font></p>
        <p align="center">&nbsp;</p>
        <p align="center"><font color="#330099" face="Verdana" size="3"><a href="selec_valida_status1.asp">Em
        Elaboração -&gt; Em Aprovação</a></font></p>
        <p align="center"><font color="#330099" face="Verdana" size="3"><a href="selec_valida_status2.asp">Validar
        Aprovação</a></font></p>
        <p align="center"><font color="#330099" face="Verdana" size="3"><a href="selec_valida_status5.asp">Em Criação
        -&gt; Criado no R/3</a></font></p>
  </form>
</body>
</html>
