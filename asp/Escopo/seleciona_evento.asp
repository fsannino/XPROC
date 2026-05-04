<%@LANGUAGE="VBSCRIPT"%> 
<%
str_Acao = request("ID")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")


str_SQl = ""
str_SQL = str_SQl & " SELECT EVEN_NR_SEQUENCIAL"
str_SQL = str_SQl & " , EVEN_DT_EVENTO"
str_SQL = str_SQl & " , EVEN_TX_DESCRICAO"
str_SQL = str_SQl & " FROM EVENTO"
str_SQL = str_SQl & " order by EVEN_DT_EVENTO "
set rds_Evento=db.execute(str_SQL)

if str_Acao = "A" then
   str_Titulo = "ALTERAÇÃO DE EVENTO"
else
   str_Titulo = "EXCLUSÃO DE EVENTO"
end if
%> 
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function Confirma()
   {
   if(document.frm1.selEvento.selectedIndex == -1)
     {
     alert("É obrigatória a seleção de um Evento!");
     document.frm1.selEvento.focus();
     return;
     } 
   else
     {	 
      if(document.frm1.txtAcao.value == "A")
        {
        document.frm1.action="alterar_evento.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
      if(document.frm1.txtAcao.value == "E")
        {
        document.frm1.action="valida_cadastro_evento.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
     }
   }
</script>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
    <td colspan="3" height="20"><table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table> </td>
  </tr>
</table>
<form name="frm1" method="post" action="">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="24%">&nbsp;</td>
    <td width="50%">&nbsp;</td>
    <td width="26%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="24%">&nbsp;</td>
    <td width="50%"><div align="center"><font size="2"><%=str_Titulo%> </font></div></td>
    <td width="26%">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><input name="txtAcao" type="hidden" id="txtAcao" value="<%=str_Acao%>"></td>
    <td><div align="center"><font color="#0000FF" size="2" face="Verdana">&nbsp;Selecione um Evento</font></div></td>
    <td>&nbsp;</td>
  </tr>
</table>

  <div align="center">
    <select name="selEvento" size="10">
      <% do while not rds_Evento.EOF 
	      str_Data = Right("00" & Day(rds_Evento("EVEN_DT_EVENTO")),2) & "/" & Right("00" & Month(rds_Evento("EVEN_DT_EVENTO")),2) & "/" & Year(rds_Evento("EVEN_DT_EVENTO"))
	%>
      <option value="<%=rds_Evento("EVEN_NR_SEQUENCIAL")%>"><%=str_Data%> - <%=rds_Evento("EVEN_TX_DESCRICAO")%></option>
      <%  rds_Evento.movenext
	Loop 
	%>
    </select>
  </div>
</form>
</body>
</html>
