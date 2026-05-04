<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

Dim vet_Dia(31)
Dim vet_Mes(12)
Dim vet_Ano(4)

For i = 1 to 31
   vet_Dia(i) = Right("00" & i,2)
next
For i = 1 to 12
   vet_Mes(i) = Right("00" & i,2)
next

vet_Ano(1) = "2001"
vet_Ano(2) = "2002"
vet_Ano(3) = "2003"
vet_Ano(4) = "2004"

str_Acao = request("txtAcao")
str_selEvento = request("selEvento")

str_SQL = ""
str_SQL = str_SQL & " SELECT "
str_SQL = str_SQL & "   EVEN_NR_SEQUENCIAL"
str_SQL = str_SQL & " , EVEN_DT_EVENTO "
str_SQL = str_SQL & " , EVEN_TX_DESCRICAO"
str_SQL = str_SQL & " FROM EVENTO"
str_SQL = str_SQL & " WHERE EVEN_NR_SEQUENCIAL =" & str_selEvento
set rs_Evento=db.execute(str_SQL)
if rs_Evento.EOF then
   response.Redirect(envia_msg_tela.asp)
end if
str_Dia_Evento = Right("00" & Day(rs_Evento("EVEN_DT_EVENTO")),2)
str_Mes_Evento = Right("00" & Month(rs_Evento("EVEN_DT_EVENTO")),2)
str_Ano_Evento = Right("00" & Year(rs_Evento("EVEN_DT_EVENTO")),4)
str_Tx_Desc = rs_Evento("EVEN_TX_DESCRICAO")
%>
<html>
<head>
<title>Altera&ccedil;&atilde;o de Eventos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function pega_tamanho()
  {
  valor=document.frm1.txtDesc.value.length;
  document.frm1.txttamanho.value=valor
  if (valor > 150) 
     {
	 str1=document.frm1.txtDesc.value;
	 str2=str1.slice(0,150);
	 document.frm1.txtDesc.value=str2;
	 valor=str2.length;
	 document.frm1.txttamanho.value=valor;
     }
  }
function manda()
{
//alert('altera_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value)
//+'&selSubModulo='+
//alert(document.frm1.selSubModulo.value)
window.location.href='altera_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value
}

function Confirma()
   {
   if(document.frm1.selDiaDtEvento.selectedIndex == 0)
     {
     alert("É obrigatória a seleção de um Dia!");
     document.frm1.selDiaDtEvento.focus();
     return;
     }
   if(document.frm1.selMesDtEvento.selectedIndex == 0)
     {
     alert("É obrigatória a seleção de um Mes!");
     document.frm1.selMesDtEvento.focus();
     return;
     }
   if(document.frm1.selAnoDtEvento.selectedIndex == 0)
     {
     alert("É obrigatória a seleção de um Ano!");
     document.frm1.selAnoDtEvento.focus();
     return;
     }
   if(document.frm1.txtDesc.value == "")
     {
     alert("É obrigatório o preenchimento da Descrição!");
     document.frm1.txtDesc.focus();
     return;
     }
     else
     {
     document.frm1.submit();
     }
   }
</script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a> 
            </div></td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div></td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div></td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div></td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div></td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div></td>
        </tr>
      </table></td>
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="24%">&nbsp;</td>
    <td width="50%">&nbsp;</td>
    <td width="26%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="24%">&nbsp;</td>
    <td width="50%"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Altera&ccedil;&atilde;o 
        de Eventos</font></div></td>
    <td width="26%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><font size="1" face="Verdana">&nbsp;</font></td>
    <td>&nbsp;</td>
  </tr>
</table>
<form name="frm1" method="post" action="valida_cadastro_evento.asp">
  <table width="100%" border="0">
    <tr> 
      <td width="27%"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><strong>Data 
          do evento: </strong></font></div></td>
      <td width="47%"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <select name="selDiaDtEvento">
          <option value="" selected>dia</option>
		  <% For j = 1 to 31 
		         if Trim(str_Dia_Evento) = Trim(vet_Dia(j)) then
				    str_Seleciona = " selected"
			     else
				    str_Seleciona = " "
				 end if		
		  %>		  
          <option value="<%=vet_Dia(j)%>" <%=str_Seleciona%>><%=vet_Dia(j)%></option>
		  <% next %>
        </select>
        / 
        <select name="selMesDtEvento">
          <option value="" selected>mes</option>
		  <% For j = 1 to 12 
		         if Trim(str_Mes_Evento) = Trim(vet_Mes(j)) then
				    str_Seleciona = " selected"
			     else
				    str_Seleciona = " "
				 end if		
		  %>		  
          <option value="<%=vet_Mes(j)%>" <%=str_Seleciona%>><%=vet_Mes(j)%></option>
		  <% next %>
        </select>
        / 
        <select name="selAnoDtEvento">
          <option value="" selected>ano</option>
		  <% For j = 1 to 4 
		         if Trim(str_Ano_Evento) = Trim(vet_Ano(j)) then
				    str_Seleciona = " selected"
			     else
				    str_Seleciona = " "
				 end if		
		  %>		  
          <option value="<%=vet_Ano(j)%>" <%=str_Seleciona%>><%=vet_Ano(j)%></option>
		  <% next %>
        </select>
        </font></td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr>
      <td><input name="txtAcao" type="hidden" id="txtAcao" value="<%=str_Acao%>">
        <input name="SelEvento" type="hidden" id="SelEvento" value="<%=str_selEvento%>"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><strong>Descri&ccedil;&atilde;o</strong></font> 
          : </div></td>
      <td><textarea name="txtDesc" cols="50" id="txtDesc" onKeyDown="javascript:pega_tamanho()"><%=str_Tx_Desc%></textarea></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font face="Verdana" size="2" color="#330099"><font size="1">Caracteres 
        digitados</font><b>&nbsp; 
        <input type="text" name="txttamanho" size="5" value="0" maxlength="50">
        </b></font><font face="Verdana" color="#330099" size="1">(Máximo 150 caracteres)</font> 
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
