<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conecta.asp" -->
<%
server.scripttimeout=99999999
response.buffer=false

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

cod= request("Registro")

Session("Registro_atual") = cod

ssql=""
ssql="SELECT *"
ssql=ssql+" FROM BACKLOG WHERE BALO_CD_COD_BACKLOG=" & COD
ssql=ssql+" ORDER BY BALO_TX_TITULO"

set rs = db.execute(ssql)

%>
<html>
<head>
<title>Carregando Backlog...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<form name="frm1" method="post" action="edita_backlog.asp">

  <p> 
    <input type="hidden" name="selMega" value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>">
    <input type="hidden" name="selModulo" value="<%=rs("SUMO_NR_CD_SEQUENCIA")%>">
    <%if len(rs("ORME_CD_ORG_MENOR"))=2 then%>
    <input type="hidden" name="Str01" value="<%=rs("ORME_CD_ORG_MENOR")%>">
    <%else%>
    <input type="hidden" name="Str01" value="<%=left(rs("ORME_CD_ORG_MENOR"),2)%>">
    <input type="hidden" name="Str02" value="<%=left(rs("ORME_CD_ORG_MENOR"),7)%>">
    <input type="hidden" name="Str03" value="<%=rs("ORME_CD_ORG_MENOR")%>">
    <%end if%>
  </p>
  <p> 
    <input type="hidden" name="txtTitulo" value="<%=rs("BALO_TX_TITULO")%>" size="150">
  </p>
  <p> 
    <input type="hidden" name="txtDescricao" value="<%=rs("BALO_TX_DESCRICAO")%>" maxlength="1000" size="100">
  </p>
  <p> 
    <input type="hidden" name="txtSolicitante" value="<%=rs("BALO_TX_SOLICITANTE")%>">
    <input type="hidden" name="txtChave" value="<%=rs("BALO_TX_CHAVE")%>">
    <input type="hidden" name="txtFone" value="<%=rs("BALO_TX_TELEFONE")%>">
  </p>
  <p> 
    <input type="hidden" name="selResponsavel" value="<%=rs("BALO_CD_RESPONSAVEL")%>">
  </p>
  <p> 
    <input type="hidden" name="selPrioridade" value="<%=rs("BALO_CD_PRIORIDADE")%>">
  </p>
  <p> 
    <input type="hidden" name="selTipo" value="<%=rs("BALO_CD_TIPO")%>">
  </p>
  <p> 
    <input type="hidden" name="selLegado" value="<%=rs("BALO_CD_LEGADO")%>">
  </p>
</form>

</body>
</html>

<script>
{
document.frm1.submit();
}
</script>
