<!--#include file="../../asp/protege/protege.asp" -->
<%
str_cenario=request("txtCenario")
str_tipo=request("txtTipo")
str_desenv=request("txtDesenv")
str_teste=request("txtTeste")
str_conf=request("txtConf")

'response.Write(" cena ")
'response.Write(request("txtCenario"))
'response.Write(" tipo ")
'response.Write(request("txtTipo"))
'response.Write(" desen ")
'response.Write(request("txtDesenv"))
'response.Write(" teste ")
'response.Write(request("txtTeste"))
'response.Write(" conf ")
'response.Write(request("txtConf"))

str_MegaProcesso=request("txtmega")
str_MegaProcesso=request("txtproc")
str_SubProcesso=request("txtsub")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

desen=0
nmuda=0

select case str_tipo

case 1
	if str_desenv=1 then
	end if 
	if  str_conf=1 then
	end if 
	if  str_teste=1 then
	end if 	
	if  desen=0 then
	end if 	
	if  nmuda=0 then
	end if 

	if str_desenv=1 and str_conf=1 and str_teste=1 and desen=0 and nmuda=0 then
		str_status="PT"
	else
		str_status="DS"
	end if
case 2
	str_desenv=0
	if str_conf=1 and str_teste=1 and desen=0 and nmuda=0 then
		str_status="PT"
	else
		str_status="DS"
end if

end select

verifica="O Status do Cenário foi alterado com Sucesso!"
cor="#330099"

if request("option")=1 then
	str_status="EE"
	str_tipo=0
	str_desenv=0
	str_conf=0
	str_teste=0
end if

ssql=""
ssql=" UPDATE " & Session("PREFIXO") & "CENARIO"
ssql=ssql+" SET CENA_TX_SITUACAO='" & str_status & "', "
ssql=ssql+" CENA_TX_SITU_DESENHO_TIPO='" & str_tipo & "', "
ssql=ssql+" CENA_TX_SITU_DESENHO_DESE='" & str_desenv & "', "
ssql=ssql+" CENA_TX_SITU_DESENHO_TESTE='" & str_teste & "', "
ssql=ssql+" ATUA_CD_NR_USUARIO='" & Session("CDUsuario") & "', "
ssql=ssql+" ATUA_DT_ATUALIZACAO= GETDATE() , "
ssql=ssql+" CENA_TX_SITU_DESENHO_CONF='" & str_conf & "' "
ssql=ssql+" WHERE CENA_CD_CENARIO='" & str_cenario & "'"

'response.write ssql

DB.EXECUTE(SSQL)
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
<form name="frm1" method="post" action="alterar_cenario.asp">
  <input type="hidden" name="INC" size="20" value="1"><input type="hidden" name="txtOpc" value="1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
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
  <table border="0" width="100%">
    <tr>
      <td width="100%" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" colspan="3">
        <p align="center"><font face="Verdana" color="#330099" size="3">Alteração 
          de Status de Cenário - <font color="#330099">Questionário</font> - <b><%=str_cenario%></b></font></p>
      </td>
    </tr>
    <tr>
      <td width="100%" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="72%" colspan="2">
        <b><font face="Verdana" color=<%=cor%> size="2"><%=verifica%></font></b>
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="72%" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="72%" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="12%">
        <p align="right"><a href="../../indexA.asp"><img src="selecao_F02.gif" width="22" height="20" border="0"></a>
      </td>
      <td width="60%">
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela Principal</font>
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="12%">
      <p align="right"><a href="gerencia_cenario_transa.asp?selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>&selCenario=<%=str_Cenario%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a>
      </td>
      <td width="60%">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
      para a tela de Edição de Cenário</font>
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="12%">
      </td>
      <td width="60%">
      </td>
    </tr>
    <tr>
      <td width="28%"></td>
      <td width="72%" colspan="2"></td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="72%" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="100%" colspan="3"></td>
    </tr>
  </table>
  </form>
<p></p>
</body>
</html>
