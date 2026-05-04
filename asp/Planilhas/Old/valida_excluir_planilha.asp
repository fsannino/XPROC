<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

linhas=request("linhas")

select case request("planilha")
case 1
	valor="IMPOSTOS - COMPRAS"
	planil="" & Session("PREFIXO") & "IMPOSTOS_COMPRAS"
	campo="[CENARIO FISCAL]"
case 2
	valor="IMPOSTOS - ARMAZENAGEM E EMPRÉSTIMOS"
	planil="" & Session("PREFIXO") & "IMPOSTOS_EMP"
	campo="[CENARIO FISCAL]"
case 3
	valor="IMPOSTOS - TRANSFERÊNCIAS"
	planil="" & Session("PREFIXO") & "IMPOSTOS_TRANSF"
	campo="[CENARIO FISCAL]"
case 4
	valor="IMPOSTOS - VENDAS"
	planil="" & Session("PREFIXO") & "IMPOSTOS_VENDAS"
	campo="[CENARIO FISCAL]"
case 5
	valor="IMPOSTOS SERVIÇOS - COMPRAS"
	planil="" & Session("PREFIXO") & "PARAM_COMPRAS"
	campo="CENARIO"
case 6
	valor="IMPOSTOS SERVIÇOS - VENDAS"
	planil="" & Session("PREFIXO") & "PARAM_VENDAS"
	campo="CENARIO"
end select

set rs=db.execute("SELECT MAX(" & campo &")AS CODIGO FROM " & planil)
maximo=rs("codigo")

contador=1
apagou=0

do until contador=maximo+1

	valor_ = request("linha_"& contador)
	select case valor_
	case "ON"
		ssql=""
		ssql="DELETE FROM " & planil & " WHERE " & campo & "=" & contador
		db.execute ssql
		
		'call grava_log(contador,planil,"D",1)
		
		apagou=apagou+1
	end select
	contador=contador+1
	
loop

if apagou<>0 then
	stat="Os Registros selecionados foram excluídos com Sucesso"
	cor="#330099"
else
	stat="Você deve selecionar as linhas que deseja excluir"
	cor="red"
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Exibição de Dados</title>
</head>

<body topmargin="0" leftmargin="0">
<form action="" method="POST" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0" width="19" height="20"></a>&nbsp;</div>
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
          <td width="21"></td>
          <td width="217"></td>
            <td width="18"></td>  <td width="16"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p align="center">
<font face="Verdana" size="3" color="#330099">Exclusão
de Linhas</font>
<p align="center"><b><font face="Arial Narrow" size="4"><%=valor%></font></b>
<p align="center">&nbsp;</p>
<p align="center"><font face="Verdana" size="2" color="<%=cor%>"><b><%=stat%></b></font></p>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
<input type="hidden" name="linhas" size="20" value="<%=contador_linha%>">
<input type="hidden" name="planilha" size="20" value="<%=request("ID")%>">
</form>
</body>

</html>
