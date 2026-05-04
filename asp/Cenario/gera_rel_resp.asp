<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql=""
ssql="select * from cenario "
ssql=ssql+"where mepr_cd_mega_processo=" & request("selMegaProcesso")
ssql=ssql+" and (cena_tx_responsavel is null or cena_tx_responsavel ='') order by cena_cd_cenario"

set rs=db.execute(ssql)

set mega=db.execute("SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso"))

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
.style4 {
	font-size: x-small;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color: #31009C;
	font-weight: bold;
}
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #31009C; font-size: x-small;}
.style12 {font-family: Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; font-size: x-small; }
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
<form name="frm1" method="post" action="">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
<tr>
   <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
           </td>
         </tr>
      </table>
    </td>
</tr>
<tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26">&nbsp;</td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27">&nbsp;</td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <p class="style4">Rela&ccedil;&atilde;o de Cen&aacute;rios sem Respons&aacute;vel</p>
  <p class="style5">Mega-Processo : <%=mega("mepr_tx_desc_mega_processo")%> </p>
  <table width="743" border="0">
    <tr bgcolor="#31009C">
      <td width="196"><span class="style12">Cen&aacute;rio</span></td>
      <td width="531"><span class="style12">Descri&ccedil;&atilde;o do Cen&aacute;rio </span></td>
    </tr>
    <%
	conta=0
	do until rs.eof=true
	%>
	<tr>
      <td><span class="style5"><%=rs("cena_cd_cenario")%></span></td>
      <td><span class="style5"><%=rs("cena_tx_titulo_cenario")%></span></td>
	</tr>
	<%
	conta=conta+1
	rs.movenext
	loop
	%>
  </table>
<p class="style5">Total de Cen&aacute;rios sem Respons&aacute;vel : <b><%=conta%></b></p>
<p>&nbsp;</p>
</form>
<p>&nbsp;</p>
</body>
</html>
