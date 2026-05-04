<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

set rs=db.execute(request("Tsql"))
int_Tot_Atualizado = 0
do until rs.eof=true
	intTipoVal = request("sel_" & rs("FUNE_CD_FUNCAO_NEGOCIO"))
	
	'*** TESTE, POIS SÓ FARÁ OS QUE FOREM DIFERENTES
	if trim(rs("FUNE_TX_TIPO_CLASS")) <> trim(intTipoVal) then
		ssql = ""
		ssql = ssql & "UPDATE " & Session("Prefixo") & "FUNCAO_NEGOCIO SET FUNE_TX_TIPO_CLASS = " & intTipoVal
		ssql = ssql & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & rs("FUNE_CD_FUNCAO_NEGOCIO")& "'"		
		db.execute(ssql)		
		int_Tot_Atualizado = int_Tot_Atualizado + 1
	end if
	rs.movenext
loop
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0">
<form method="POST" action="" name="frm1">          
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

 
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%"> 
      <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Libera Fun&ccedil;&atilde;o para mapeamento </font></td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Total atualizado </font>: <%=int_Tot_Atualizado%></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
<table border="0" width="94%" height="114">
  <tr>
    <td width="34%" height="42"></td>
    <td width="5%" align="center" height="42">
      <p align="center"><a href="../../indexA.asp"><img src="selecao_F02.gif" border="0" align="left"></a></td>
    <td width="61%" colspan="2" height="42"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta para tela Principal</font></td>
  </tr>
  <tr>
    <td width="34%" height="39"></td>
    <td width="5%" align="center" height="39"><a href="sel_func_libera_mapeamento.asp?pOpt=2"><img src="selecao_F02.gif" border="0" align="left"></a></td>
    <td width="61%" colspan="2" height="39"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar para a tela de Sele&ccedil;&atilde;o de Classificação para os Cursos</font></td>
  </tr>
  <tr>
    <td width="34%" height="21"></td>
    <td width="5%" align="center" height="21"></td>
    <td width="36%" height="21"></td>
    <td width="25%" height="21"></td>
  </tr>
</table>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
</form>

<p>&nbsp;</p>

</body>

</html>
