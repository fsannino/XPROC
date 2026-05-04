<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

select case request("ID")
case 1
	valor="IMPOSTOS - COMPRAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "IMPOSTOS_COMPRAS ORDER BY [CENARIO FISCAL]")
case 2
	valor="IMPOSTOS - ARMAZENAGEM E EMPRÉSTIMOS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "IMPOSTOS_EMP ORDER BY [CENARIO FISCAL]")
case 3
	valor="IMPOSTOS - TRANSFERÊNCIAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "IMPOSTOS_TRANSF ORDER BY [CENARIO FISCAL]")
case 4
	valor="IMPOSTOS - VENDAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "IMPOSTOS_VENDAS ORDER BY [CENARIO FISCAL]")
case 5
	valor="IMPOSTOS SERVIÇOS - COMPRAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "PARAM_COMPRAS ORDER BY CENARIO")
case 6
	valor="IMPOSTOS SERVIÇOS - VENDAS"
	Set rs = db.Execute("SELECT * FROM " & Session("PREFIXO") & "PARAM_VENDAS ORDER BY CENARIO")
end select

quantos = rs.Fields.Count

contador=0
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Exibição de Dados</title>
</head>

<body topmargin="0" leftmargin="0">
<form method="post" action="valida_altera_planilha.asp">
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
            <td width="26"><img border="0" src="../../imagens/confirma_f02.gif" width="24" height="24" onclick="javascript:submit()"></td>
          <td width="119"><b><font size="2" face="Verdana" color="#330099">Alterar</font></b></td>
          <td width="27">&nbsp;</td>
          <td width="80"></td>
          <td width="88"></td>
            <td width="18"></td>  <td width="16"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p><b><font face="Arial Narrow" size="4"><%=valor%>  - </font></b><font face="Arial Narrow" size="3"> 
<%IF RS.EOF=FALSE THEN%>
<%END IF%><input type="hidden" name="planilha" size="20" value="<%=REQUEST("ID")%>">Altere
os dados desejados e clique em &quot;Alterar&quot;</font></p>
<b><font face="Arial Narrow" size="4">
<table border="0" width="100%">
   <tr>
    <%do until contador=quantos%>
    <td width="100%" align="center" bgcolor="#AFCDD8"><font face="Arial Narrow" size="2"><%=ucase(rs.fields(contador).name)%></font></td>
    <%
    contador=contador+1
    loop
    CONTADOR=0
    %>
    </tr>
   <%
   contador_linha=1
	DO UNTIL RS.EOF=TRUE
	IF COR="WHITE" THEN
		COR="#E4E4E4"
	ELSE
		COR="WHITE"
	END IF
   %>
   <tr>
   <%
   do until contador=quantos%>
   <td width="56%" bgcolor="<%=cor%>">
    <textarea rows="2" name="<%=ucase(rs.fields(contador).name) & "_" & contador_linha%>" cols="20"><%=ucase(rs.fields(contador).value)%></textarea></td>
   <%
   contador = contador+1
   loop
   CONTADOR=0
   %>
  </tr>
  <%
  RS.MOVENEXT
  contador_linha=contador_linha+1
  LOOP
  %>
</table>
<input type="hidden" name="linhas" size="20" value="<%=contador_linha-1%>">
</form>
</body>

</html>
