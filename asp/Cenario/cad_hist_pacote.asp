<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

Session("Conn_String_Cogest_Gravacao")="Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs1=db.execute("SELECT DISTINCT CENA_TX_PACT_TESTE FROM CENARIO WHERE CENA_TX_PACT_TESTE<>'' ORDER BY CENA_TX_PACT_TESTE")

valversao="1"

if request("pacte")<>"" then
	set fonte = db.execute("SELECT * FROM PACOTE_TESTES WHERE PCTE_TX_PACT_TESTE='" & request("pacte") & "'")
	set maximo = db.execute("SELECT MAX(PCTE_NM_VERSAO)AS VERSAO FROM PACOTE_TESTES WHERE PCTE_TX_PACT_TESTE='" & request("pacte") & "'")
	
	valversao=maximo("VERSAO")
	
	if valversao>0 then
		valversao=valversao+1
	else
		valversao=1
	end if
	
else
	if request("excel")<>1 then
		set fonte = db.execute("SELECT * FROM PACOTE_TESTES WHERE PCTE_TX_PACT_TESTE='YYYYYY'")
	else
		set fonte = db.execute("SELECT * FROM PACOTE_TESTES ORDER BY PCTE_TX_PACT_TESTE")	
	end if
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
<script language="JavaScript">
<!--

function Confirma() 
{
if(document.frm1.selpacote.selectedIndex == 0)
{
alert("É obrigatória a seleção de um PACOTE DE TESTES!");
document.frm1.selpacote.focus();
return;
}
if(document.frm1.txtdescricao.value == "")
{
alert("É obrigatória o preenchimento da DESCRIÇÃO!");
document.frm1.txtdescricao.focus();
return;
}
if(document.frm1.txtdatainicio.value == "")
{
alert("É obrigatória a seleção da DATA DE INÍCIO DE TESTES!");
document.frm1.txtdatainicio.focus();
return;
}
if(document.frm1.txtsemana.value == "")
{
alert("É obrigatório o preenchimento das SEMANAS DE TESTES!");
document.frm1.txtsemana.focus();
return;
}
if(document.frm1.selmotivo.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MOTIVO DE ALTERAÇÃO!");
document.frm1.selmotivo.focus();
return;
}
else
{ 
document.frm1.submit();
}
}
//  End -->
</script>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
<script language="JavaScript" src="pupdate.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#CC3300" vlink="#CC3300" alink="#CC3300">
<%if request("excel")<>1 then%>
<form name="frm1" method="POST" action="valida_cad_hist_pacote.asp">
  <table width="903" height="105" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" style="margin-bottom: 0">
    <tr> 
      <td height="66" colspan="2">&nbsp;</td>
      <td height="66" colspan="2">&nbsp;</td>
      <td valign="top" colspan="2" height="66"> 
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="39" width="1">&nbsp; </td>
      <td height="39" width="1">&nbsp;</td>
      <td height="39" width="625"> 
        <table width="625" border="0" align="center">
          <tr> 
            <td width="26" height="30"></td>
            <td width="104"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="24"><a href="javascript:Confirma()"><img src="confirma_f02.gif" width="24" height="24" border="0"></a></td>
            <td width="143"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Cadastrar</font></b></font></td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="78"></td>
            <td width="28"><div align="right"><a href="cad_hist_pacote.asp?excel=1" target="_blank"><img src="../Apoio/excel.jpg" width="27" height="24" border="0"></a></div></td>
            <td width="107"><font color="#330099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Exportar 
              Relat&oacute;rio para o Excel</b></font></td>
          </tr>
        </table></td>
      <td colspan="2" height="39"> </td>
      <td height="39" width="274"></td>
    </tr>
  </table>
  <table width="90%" border="0">
    <tr> 
      <td width="21%"></td>
      <td width="55%"></td>
      <td width="24%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="center"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif">Cadastro 
          de Hist&oacute;rico de Pacote de Testes</font></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="74%" border="0" align="center">
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="24%"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Pacote 
        de Testes</font></td>
      <td width="26%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <select name="selpacote" id="selpacote" onChange="window.location='cad_hist_pacote.asp?pacte='+this.value">
          <option value="xxx">== Selecione um Pacote ==</option>
          <%
		  do until rs1.eof=true
			if request("pacte")=rs1("CENA_TX_PACT_TESTE") then
				ol1="selected"
			else
				ol1=""
			end if
		  %>
          <option <%=ol1%> value="<%=rs1("CENA_TX_PACT_TESTE")%>"><%=rs1("CENA_TX_PACT_TESTE")%></option>
          <%
		  rs1.movenext
		  loop
		  %>
        </select>
        </font></td>
      <td width="20%" align="center" valign="middle"> <div align="center"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Vers&atilde;o</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <input name="txtversao" type="text" id="txtversao" size="5" maxlength="3" readonly value="<%=valversao%>">
          </font></div></td>
      <td width="29%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></td>
      <td colspan="3" rowspan="2" valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <textarea name="txtdescricao" cols="60" rows="2" id="txtdescricao"></textarea>
        </font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="43">&nbsp;</td>
      <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Data 
        de In&iacute;cio de Testes</font></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="txtdatainicio" type="text" id="txtdatainicio" size="30" maxlength="10" readonly>
        </font></td>
      <td align="left" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<img src="../../imagens/calendario3.gif" alt="Clique aqui para selecionar a Data de In&iacute;cio dos Testes" width="37" height="34" align="absmiddle" onClick="getCalendarFor(document.frm1.txtdatainicio)"></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Semanas 
        de Testes</font></td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <input name="txtsemana" type="text" id="txtsemana" size="10" maxlength="3">
        </font></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Motivo 
        da Altera&ccedil;&atilde;o </font></td>
      <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <select name="selmotivo" id="selmotivo">
          <option value="XXX">== Selecione o Motivo ==</option>
          <option value="ATRASO DESENVOLVIMENTO">ATRASO DESENVOLVIMENTO</option>
          <option value="ATRASO CONFIGURAÇÃO">ATRASO CONFIGURAÇÃO</option>
          <option value="FALTA DE RECURSOS">FALTA DE RECURSOS</option>
          <option value="OUTROS">OUTROS</option>
        </select>
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es</font></td>
      <td colspan="3" rowspan="2" valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <textarea name="txtobs" cols="60" rows="2" id="txtobs"></textarea>
        </font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <%end if%>
  <%
  if fonte.eof=false then
  %>
  <p align="center"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Hist&oacute;rico 
    Atual de Cadastro</strong></font></p>
   <table width="98%" border="0" bordercolor="#333333">
    <tr bgcolor="#330099"> 
      <td width="13%" height="24"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Pacote</font></strong></td>
      <td width="6%"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Vers&atilde;o</font></strong></td>
      <td width="21%"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></strong></td>
      <td width="16%"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
        de In&iacute;cio de Testes</font></strong></td>
      <td width="12%"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Semanas 
        de Teste</font></strong></td>
      <td width="14%"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Motivo 
        da Altera&ccedil;&atilde;o</font></strong></td>
      <td width="18%"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Observa&ccedil;&otilde;es</font></strong></td>
    </tr>
	<%
	do until fonte.eof=true
	%>
    <tr> 
      <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=fonte.fields(0).value%></font></td>
      <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=fonte.fields(1).value%></font></td>
      <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=fonte.fields(2).value%></font></td>
	  <%
		datai=day(fonte.fields(3).value) &"/"& month(fonte.fields(3).value) & "/" & year(fonte.fields(3).value)	  
	  %>
      <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=datai%></font></td>
      <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=fonte.fields(4).value%></font></td>
      <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=fonte.fields(5).value%></font></td>
      <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=fonte.fields(6).value%></font></td>
    </tr>
	<%
	fonte.movenext
	loop
	%>
  </table>
  </form>
  <%end if%>
<p></p>
<%if request("excel")<>1 then%>
<!-- PopUp Calendar BEGIN -->
<script language="JavaScript">
if (document.all) {
 document.writeln("<div id=\"PopUpCalendar\" style=\"position:absolute; left:0px; top:0px; z-index:7; width:200px; height:77px; overflow: visible; visibility: hidden; background-color: #FFFFFF; border: 1px none #000000\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout(\'hideCalendar()\',500)\">");
 document.writeln("<div id=\"monthSelector\" style=\"position:absolute; left:0px; top:0px; z-index:9; width:181px; height:27px; overflow: visible; visibility:inherit\">");}
else if (document.layers) {
 document.writeln("<layer id=\"PopUpCalendar\" pagex=\"0\" pagey=\"0\" width=\"200\" height=\"200\" z-index=\"100\" visibility=\"hide\" bgcolor=\"#FFFFFF\" onMouseOver=\"if(ppcTI){clearTimeout(ppcTI);ppcTI=false;}\" onMouseOut=\"ppcTI=setTimeout('hideCalendar()',500)\">");
 document.writeln("<layer id=\"monthSelector\" left=\"0\" top=\"0\" width=\"181\" height=\"27\" z-index=\"9\" visibility=\"inherit\">");}
else {
 document.writeln("<p><font color=\"#FF0000\"><b>Error ! The current browser is either too old or too modern (usind DOM document structure).</b></font></p>");}
</script>
<noscript></noscript>
<table border="1" cellspacing="1" cellpadding="2" width="200" bordercolorlight="#000000" bordercolordark="#000000" vspace="0" hspace="0"><form name="ppcMonthList"><tr><td align="center" bgcolor="#CCCCCC"><a href="javascript:moveMonth('Back')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b>&lt;&nbsp;</b></font></a><font face="MS Sans Serif, sans-serif" size="1"> 
<select name="sItem" onMouseOut="if(ppcIE){window.event.cancelBubble = true;}" onChange="switchMonth(this.options[this.selectedIndex].value)" style="font-family: 'MS Sans Serif', sans-serif; font-size: 9pt"><option value="0" selected>2000
  • Janeiro</option><option value="1">2000 • Fevereiro</option><option value="2">2000
  • Março</option><option value="3">2000 • Abril</option><option value="4">2000
  • Maio</option><option value="5">2000 • Junho</option><option value="6">2000
  • Julho</option><option value="7">2000 • Agosto</option><option value="8">2000
  • Setembro</option><option value="9">2000 • Outubro</option><option value="10">2000
  • Novembro</option><option value="11">2000 • Dezembro</option><option value="0">2001
  • Janeiro</option></select></font><a href="javascript:moveMonth('Forward')" onMouseOver="window.status=' ';return true;"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000"><b>&nbsp;&gt;</b></font></a></td></tr></form></table>
<table border="1" cellspacing="1" cellpadding="2" bordercolorlight="#000000" bordercolordark="#000000" width="200" vspace="0" hspace="0"><tr align="center" bgcolor="#CCCCCC"><td width="20" bgcolor="#FFFFCC"><b><font face="Arial" size="1">Dom</font></b></td><td width="20"><b><font face="Arial" size="1">Seg</font></b></td><td width="20"><b><font face="Arial" size="1">Ter</font></b></td><td width="20"><b><font face="Arial" size="1">Qua</font></b></td><td width="20"><b><font face="Arial" size="1">Qui</font></b></td><td width="20"><b><font face="Arial" size="1">Sex</font></b></td><td width="20" bgcolor="#FFFFCC"><b><font face="Arial" size="1">Sab</font></b></td></tr></table>
<script language="JavaScript">
if (document.all) {
 document.writeln("</div>");
 document.writeln("<div id=\"monthDays\" style=\"position:absolute; left:0px; top:52px; z-index:8; width:200px; height:17px; overflow: visible; visibility:inherit; background-color: #FFFFFF; border: 1px none #000000\">&nbsp;</div></div>");}
else if (document.layers) {
 document.writeln("</layer>");
 document.writeln("<layer id=\"monthDays\" left=\"0\" top=\"52\" width=\"200\" height=\"17\" z-index=\"8\" bgcolor=\"#FFFFFF\" visibility=\"inherit\">&nbsp;</layer></layer>");}
else {/*NOP*/}
</script>
<!-- PopUp Calendar END -->
<%end if%>
</body>
</html>