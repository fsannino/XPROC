 
<%
cenario=request("ID")
opt=request("OPTION")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='"& trim(cenario) & "'"
set rs=db.execute(ssql)

valor=rs("CENA_TX_TITULO_CENARIO")

set rstrans=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='"& trim(cenario) & "'")
%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">

function fechar()
{
window.close();
}

function mover()
{
window.moveTo(100	,200)
}

</script>

<body topmargin="0" leftmargin="0" onload="javascript:mover()">
<table border="0" width="90%">
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2"></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="83%"><font face="Verdana" size="2" color="#000080"><b>Relação
      de Transações Existentes</b></font></td>
    <td width="5%">
    </td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2"></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2"></td>
  </tr>
  <tr>
    <td width="12%"><font face="Verdana" size="2" color="#000080"><b>Cenário : </b></font></td>
    <td width="88%" colspan="2"><font face="Verdana" size="2" color="#000080"><%=cenario%>-<%=valor%></font></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2"></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2"></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2"><font face="Verdana" size="2"><select size="8" name="D1" multiple>
    	<%if rstrans.eof=true then%>
      <option>Nenhuma Transação Cadastrada para este Cenário</option>
		<%end if%>
		<%DO UNTIL RSTRANS.EOF=TRUE
	   IF NOT ISNULL(rstrans("MEPR_CD_MEGA_PROCESSO")) THEN
	   SSQL1="SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rstrans("MEPR_CD_MEGA_PROCESSO")
	   SET RS1=DB.EXECUTE(SSQL1)
	   VALOR2=RS1("MEPR_TX_ABREVIA")
	   else
	   valor2="**"
	   END IF
	   %>
      <option><%=VALOR2%> - <%=RSTRANS("TRAN_CD_TRANSACAO")%> - <%=left(RSTRANS("CETR_TX_DESC_TRANSACAO"),28)%></option>
      <%
      RSTRANS.MOVENEXT
      LOOP
      %>
      </select></font></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2"></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2"></td>
  </tr>
  <tr>
    <td width="12%"></td>
    <td width="88%" colspan="2">
      <p align="right"><b><a href="#" onclick="javascript:fechar()"><font face="Verdana" size="2" color="#800000">Fechar
      Janela</font></a></b></td>
  </tr>
</table>

</body>

</html>
