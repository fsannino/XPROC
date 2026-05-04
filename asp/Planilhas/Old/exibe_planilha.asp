<%
set db = Server.CreateObject("ADODB.Connection")

ID=request("ID")

with db
	.Provider="Microsoft.Jet.Oledb.4.0"
	.Properties("Extended Properties").value="Excel 8.0"
	.Open server.mappath("../planilhas/plans/" & ID & ".xls")
end with

Set rs = db.Execute("SELECT * FROM [plan1$]")

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

<SCRIPT>
function excluir()
{
window.location.href="excluir_planilha.asp?ID="+this.planilha.value
}

function altera()
{
window.location.href="altera_planilha.asp?ID="+this.planilha.value
}

function inclui()
{
window.location.href="incluir_planilha.asp?ID="+this.planilha.value
}

</SCRIPT>

<body topmargin="0" leftmargin="0">

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
<p><b><font face="Arial Narrow" size="4"><%=valor%>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</font></b><input type="button" value="Incluir " name="B3" onclick="javascript:inclui()">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<%IF RS.EOF=FALSE THEN%>
<input type="submit" value="Alterar" name="B1" onclick="javascript:altera()">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" value="Excluir" name="B2" onclick="javascript:excluir()">
 <%END IF%><input type="hidden" name="planilha" size="20" value="<%=REQUEST("ID")%>"></p>
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
    <p align="center"><font face="Arial Narrow" size="2"><%=ucase(rs.fields(contador).value)%></FONT></td>
   <%
   contador = contador+1
   loop
   CONTADOR=0
   %>
  </tr>
  <%
  RS.MOVENEXT
  LOOP
  %>
</table>

</body>

</html>
