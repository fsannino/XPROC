 
<!--#include file="../../asp/protege/protege.asp" -->
<%
publico=request("txtpublico")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

atual=0

set rs=db.execute("SELECT MAX(TPPP_CD_TIPO_PUB_PRINCIPAL)AS CODIGO FROM " & Session("PREFIXO") & "TIPO_PUBLICO_PRINCIPAL")

ATUAL=RS("CODIGO")

if atual=0 then
	atual=1
else
	atual=atual+1
end if

codigo=ATUAL

ssql=""
ssql="INSERT INTO " & Session("PREFIXO") & "TIPO_PUBLICO_PRINCIPAL "
ssql=ssql & "VALUES('" & ucase(publico) & "', "
ssql=ssql+"'" & codigo & "', "
ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE())"


db.execute(ssql)
'call grava_log(codigo,"TIPO_PUBLICO_PRINCIPAL","I",1)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body topmargin="0" leftmargin="0">
<form method="POST" action="" name="frm1">
<input type="hidden" name="txtpub" size="20"><input type="hidden" name="txtQua" size="20">
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
          <td width="26"></td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Cadastro
        de Público Principal / Depto</font></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2">O
Registro foi incluído com sucesso com o </font><font face="Verdana" color="#330099" size="2"> Código
</font><font face="Verdana" color="#330099" size="3"> <%=codigo%></font></b></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="100%">
  <tr>
    <td width="33%"></td>
            <td width="48"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
  </tr>
  <tr>
    <td width="33%"></td>
            <td width="48">
              <p align="right"><a href="cad_funcao.asp"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar
              para a tela de Cadastro de Fun&ccedil;&atilde;o R/3</font></td>
  </tr>
  <tr>
    <td width="33%"></td>
    <td width="9%"></td>
    <td width="58%"></td>
  </tr>
</table>
  </form>

<p>&nbsp;</p>

</body>

</html>
