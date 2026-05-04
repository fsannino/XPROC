<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Acao = request("ID")

%>
<html>
<head>
<title>Inclus&atilde;o de Eventos</title>
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
    <td width="50%"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Inclus&atilde;o 
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
          <option value="01">01</option>
          <option value="02">02</option>
          <option value="03">03</option>
          <option value="04">04</option>
          <option value="05">05</option>
          <option value="06">06</option>
          <option value="07">07</option>
          <option value="08">08</option>
          <option value="09">09</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23">23</option>
          <option value="24">24</option>
          <option value="25">25</option>
          <option value="26">26</option>
          <option value="27">27</option>
          <option value="28">28</option>
          <option value="29">29</option>
          <option value="30">30</option>
          <option value="31">31</option>
        </select>
        / 
        <select name="selMesDtEvento">
          <option value="" selected>mes</option>
          <option value="01">01</option>
          <option value="02">02</option>
          <option value="03">03</option>
          <option value="04">04</option>
          <option value="05">05</option>
          <option value="06">06</option>
          <option value="07">07</option>
          <option value="08">08</option>
          <option value="09">09</option>
          <option value="10">10</option>
          <option value="11">11</option>
          <option value="12">12</option>
          <option value="13">13</option>
          <option value="14">14</option>
          <option value="15">15</option>
          <option value="16">16</option>
          <option value="17">17</option>
          <option value="18">18</option>
          <option value="19">19</option>
          <option value="20">20</option>
          <option value="21">21</option>
          <option value="22">22</option>
          <option value="23">23</option>
          <option value="24">24</option>
          <option value="25">25</option>
          <option value="26">26</option>
          <option value="27">27</option>
          <option value="28">28</option>
          <option value="29">29</option>
          <option value="30">30</option>
          <option value="31">31</option>
        </select>
        / 
        <select name="selAnoDtEvento">
          <option value="" selected>ano</option>
          <option value="2000">2000</option>
          <option value="2001">2001</option>
          <option value="2002">2002</option>
          <option value="2003">2003</option>
          <option value="2004">2004</option>
          <option value="2005">2005</option>
          <option value="2006">2006</option>
        </select>
        </font></td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr>
      <td><input name="txtAcao" type="hidden" id="txtAcao" value="<%=str_Acao%>"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><strong>Descri&ccedil;&atilde;o</strong></font> 
          : </div></td>
      <td><textarea name="txtDesc" cols="50" id="txtDesc" onKeyDown="javascript:pega_tamanho()"></textarea></td>
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
