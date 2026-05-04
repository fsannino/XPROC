<%@LANGUAGE="VBSCRIPT"%> 
<%
server.scripttimeout=99999999
response.buffer=false

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

Motivo = request("txtmotivo")

ssql = request("txtQuery")
set fonte = db.execute(ssql)

i = 0
reg = fonte.RecordCount

do until i = reg
	st_atual = request(fonte("USAP_CD_USUARIO") & "_" & fonte("CURS_CD_CURSO"))
	if st_atual = 1 then
		sqla = "UPDATE USUARIO_APROVADO SET USAP_TX_APROVEITAMENTO='LM', ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "', ATUA_TX_OPERACAO='A', MOTI_NR_CD_MOTIVO= '" & Motivo & "', USAP_DT_LIBERADO_MANUAL=GETDATE() WHERE USAP_CD_USUARIO='" & fonte("USAP_CD_USUARIO") & "' AND CURS_CD_CURSO='" & fonte("CURS_CD_CURSO") & "'"
		db.execute(sqla)
	else
		sqla = "UPDATE USUARIO_APROVADO SET USAP_TX_APROVEITAMENTO='', ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "', ATUA_TX_OPERACAO='A', ATUA_DT_ATUALIZACAO=GETDATE() WHERE USAP_CD_USUARIO='" & fonte("USAP_CD_USUARIO") & "' AND CURS_CD_CURSO='" & fonte("CURS_CD_CURSO") & "'"
		db.execute(sqla)
	end if	
	i = i + 1
	fonte.movenext
loop

db.close
set db = nothing
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0">
<form method="POST" action="" name="frm1">

<input type="hidden" name="txtpub" size="20">
<input type="hidden" name="txtQua" size="20">

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr>
            <td bgcolor="#330099" width="51" valign="middle" align="right"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="49" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="50" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
        <tr>
            <td bgcolor="#330099" height="12" width="51" valign="middle" align="right"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="49" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="50" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center" height="27">
        <tr>
            <td width="26" height="23"></td>
          <td width="26" height="23"></td>
          <td width="195" height="23"></td>
            <td width="27" height="23"></td>  <td width="50" height="23"></td>
          <td width="28" height="23"></td>
          <td width="26" height="23">&nbsp;</td>
          <td width="159" height="23"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

  <table border="0" width="88%" height="33">
    <tr>
      <td width="6%" height="16"></td>
      <td width="10%" height="16"></td>
      <td width="42%" height="16"><font face="Verdana" color="#000080">Treinamento - Liberação Manual de Usuários em Cursos (<b>LM</b>)</font></td>
    </tr>
    </table>
<p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2"><%=topico%></font><font face="Verdana" color="#330099" size="3"></font></b></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="889" height="101">
  <tr>
    <td width="241" height="28">&nbsp;</td>
            <td width="34" height="28">&nbsp;</td>
            <td height="28" width="594"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#003366">Liberação Manual efetuada com Sucesso!</font></b></td>
  </tr>
  <tr>
    <td width="241" height="28">&nbsp;</td>
            <td width="34" height="28">&nbsp;</td>
            <td height="28" width="594">&nbsp;</td>
  </tr>
  <tr>
    <td width="241" height="34"></td>
            <td width="34" height="34" align="center"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
            <td height="34" width="594"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
  </tr>
  <tr>
    <td width="241" height="32"></td>
            <td width="34" height="32" align="center">
              <p align="right"><a href="seleciona_lm.asp"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
            <td height="32" width="594"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar para a tela de </font><font face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><font size="2">Liberação Manual de Usuários (<b>LM</b>)</font></font></td>
  </tr>
</table>
  </form>

<p>&nbsp;</p>

</body>

</html>