<%
Response.Buffer=false

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

StrTransacao=request("selTransacao")

set rs=db.execute("SELECT * FROM TRANSACAO ORDER BY TRAN_CD_TRANSACAO, TRAN_TX_DESC_TRANSACAO")

achou=0

if strTransacao<>"" then
		set fonte = db.execute("SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM FUN_NEG_TRANSACAO WHERE TRAN_CD_TRANSACAO='" & StrTransacao & "' ORDER BY FUNE_CD_FUNCAO_NEGOCIO")
		if fonte.eof=true then
			achou=0
		end if
	else
		set fonte = db.execute("SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM FUN_NEG_TRANSACAO WHERE TRAN_CD_TRANSACAO='SOMEBODYWANTS'")
		StrTransacao="NENHUMA"
		achou=3
end if

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function manda(v)
{
if (v!='NENHUMA')
{
	window.location="consulta_func_trans.asp?selTransacao="+v
}
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form>
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
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
<table border="0" width="73%">
  <tr>
    <td width="11%">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    <td width="89%">
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><b>Selecione a Transação Desejada :</b></font></p>
    </td>
  </tr>
  <tr>
    <td width="11%">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    <td width="89%">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><b> <select size="1" name="selTransacao" onChange="manda(this.value)">
  <option value="NENHUMA">== Selecione a Transação ==</option>
  <%do until rs.eof=true
  if trim(StrTransacao)=trim(rs("TRAN_CD_TRANSACAO"))then
  %>
  <option selected value="<%=rs("TRAN_CD_TRANSACAO")%>"><%=rs("TRAN_CD_TRANSACAO")%> - <%=LEFT((rs("TRAN_TX_DESC_TRANSACAO")),40)%></option>
  <%else%>
  <option value="<%=rs("TRAN_CD_TRANSACAO")%>"><%=rs("TRAN_CD_TRANSACAO")%> - <%=LEFT((rs("TRAN_TX_DESC_TRANSACAO")),40)%></option>
  <%  
  end if
  rs.movenext
  loop
  %>
</select></b></font></td>
  </tr>
</table>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>

<table border="0" width="722">
<%if fonte.eof=false then%>
  <tr>
    <td width="280">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    <td width="138" bgcolor="#000080">
      <b><font face="Verdana" size="2" color="#FFFFFF">Código</font></b>
    </td>
    <td width="793" bgcolor="#000080">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2" color="#FFFFFF">Função
      de Negócio Associada </font></b></p>
    </td>
  </tr>
  <%
  do until fonte.eof=true
  
  achou=1
  
  IF COR="#E2E2E2" THEN
  	COR="WHITE"
  ELSE
  	COR="#E2E2E2"
  END IF
  
  %>
  <tr>
    <td width="280">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    <td width="138" bgcolor="<%=COR%>">
      <font size="1" face="Verdana"><%=UCASE(FONTE("FUNE_CD_FUNCAO_NEGOCIO"))%></font></td>
    <%
    set temp=db.execute("SELECT * FROM FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & FONTE("FUNE_CD_FUNCAO_NEGOCIO") & "'")
    %>
    <td width="793" bgcolor="<%=COR%>">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font size="1" face="Verdana"><%=TEMP("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
  </tr>
  <%
  fonte.movenext
  loop
  %>
</table>
<%else
if achou=0 then
%>
<p><b><font color="#800000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Nenhum Registro Encontrado para a Seleção</font></b></p>
<%
else
if achou=3 then
%>
<p><b><font color="#800000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><font color="#330099">&nbsp;&nbsp;&nbsp;
</font></b><font color="#330099"><font face="Arial Narrow">Selecione a Transação Desejada na lista acima...</font></font></p>
<%
end if
end if
end if
%>
<p>&nbsp;</p>
</form>
</body>
</html>
