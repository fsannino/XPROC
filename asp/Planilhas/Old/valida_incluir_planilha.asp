<%@LANGUAGE="VBSCRIPT"%> 
 
<%
SERVER.SCRIPTTIMEOUT=99999999

ID=request("planilha")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

select case request("ID")
case 1
	valor="IMPOSTOS - COMPRAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "IMPOSTOS_COMPRAS ORDER BY [CENARIO FISCAL]")
	plan1="" & Session("PREFIXO") & "IMPOSTOS_COMPRAS"
case 2
	valor="IMPOSTOS - ARMAZENAGEM E EMPRÉSTIMOS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "IMPOSTOS_EMP ORDER BY [CENARIO FISCAL]")
	plan1="" & Session("PREFIXO") & "IMPOSTOS_EMP"
case 3
	valor="IMPOSTOS - TRANSFERÊNCIAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "IMPOSTOS_TRANSF ORDER BY [CENARIO FISCAL]")
	plan1="" & Session("PREFIXO") & "IMPOSTOS_TRANSF"
case 4
	valor="IMPOSTOS - VENDAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "IMPOSTOS_VENDAS ORDER BY [CENARIO FISCAL]")
	plan1="" & Session("PREFIXO") & "IMPOSTOS_VENDAS"
case 5
	valor="IMPOSTOS SERVIÇOS - COMPRAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "PARAM_COMPRAS ORDER BY CENARIO")
	plan1="" & Session("PREFIXO") & "PARAM_COMPRAR"
case 6
	valor="IMPOSTOS SERVIÇOS - VENDAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "PARAM_VENDAS ORDER BY CENARIO")
	plan1="" & Session("PREFIXO") & "PARAM_VENDAS"
end select

linhas=request("linhas")

VALOR_CAMPOS=RS.FIELDS.COUNT

contador=1

do until contador = linhas + 1

	CONTA_CAMPO=0
	
	ssql=""
	ssql="INSERT INTO " & plan1
	
	valor_atual=REQUEST(RS.FIELDS(CONTA_CAMPO).NAME & "_1")
	
	if len(valor_atual)=0 then
		valor_atual=0
	end if
	
	ssql=ssql&" VALUES (" & valor_atual & ","
	
	CONTA_CAMPO = CONTA_CAMPO + 1
	
	DO UNTIL CONTA_CAMPO=VALOR_CAMPOS
		ssql=ssql&"'" & (REQUEST(RS.FIELDS(CONTA_CAMPO).NAME & "_1")) & "',"
		CONTA_CAMPO = CONTA_CAMPO + 1 
	LOOP
	
	contador=contador+1
	
	tamanho=len(ssql)
	tamanho=tamanho-1
	
	ssql=left(ssql,tamanho) &")"
	
	response.write ssql
	
	db.execute(ssql)
	
	'call grava_log(valor_atual,planil,"I",1)

loop

stat="Os Registros foram incluídos com Sucesso"
cor="#330099"
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
<font face="Verdana" size="3" color="#330099">Inclusão de Linhas</font>
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
