<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("selMegaProcesso")<>0 then
	mega2=request("selMegaProcesso")
else
	mega2=0
end if

set mega=db.execute("SELECT * FROM MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set onda=db.execute("SELECT * FROM ONDA ORDER BY ONDA_TX_DESC_ONDA")

set status=db.execute("SELECT * FROM SITUACAO_GERAL WHERE SITU_TX_REFERENTE='CENARIO' ORDER BY SITU_TX_DESC_SITUACAO")

set evento=db.execute("SELECT * FROM EVENTO ORDER BY EVEN_DT_EVENTO")

if mega2<>0 then
	set assunto=db.execute("SELECT * FROM SUB_MODULO WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & mega2 & "%' ORDER BY SUMO_TX_DESC_SUB_MODULO")
else
	set assunto=db.execute("SELECT * FROM SUB_MODULO WHERE MEPR_CD_MEGA_PROCESSO_TODOS = '0' ORDER BY SUMO_TX_DESC_SUB_MODULO")
end if	
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
</script>
</head>

<script>
function Confirma()
{
var a=document.frm1.data01.value
var chk    = 0;
var maxDay = 0;

var dd = a.slice(0,2)
var mm = a.slice(3,5)
var yyyy = a.slice(6,10)

maxDay = max_day(mm, yyyy);  

if((dd <= 0) || (dd > maxDay))
{ chk = 1;}
else if((mm <= 0) || (mm > 12))
{ chk = 1;}
else if((yyyy <= 0))
{ chk = 1;} 

if(chk == 1)
{ 
alert('Data Inválida! Tente novamente');
document.frm1.data01.value='';
document.frm1.data01.focus()
}
else
{ 
document.frm1.submit();
}
}
function max_day(mn, yr)
{
   var mDay;
if((mn == 4) || (mn == 6) || (mn == 9) || (mn == 11))
{ 
mDay = 30;
}
else if(mn == 2)
{
mDay = isLeapYear(yr) ? 29 : 28;    
}
else
{
mDay = 31;
}
return mDay; 
}

function isLeapYear(yr)
{
if (yr % 2 == 0) 
return true;
return false;
}
</script>

<script>
function foca()
{
document.frm1.data01.focus();
}

function FormataData(Campo,teclapres) {
	var tam_ = event.srcElement.value
	tam=tam_.length
	if(tam<10){
		var tecla = teclapres.keyCode;
		if((tecla >= 48 && tecla <= 57) || (tecla >= 96 && tecla <= 105))
		{
			vr = event.srcElement.value;
			vr = vr.replace( ".", "" );
			vr = vr.replace( "/", "" );
			vr = vr.replace( "/", "" );
			tam = vr.length + 1;
			if ( tecla != 9 && tecla != 8 ){
				if ( tam > 2 && tam < 5 )
				{
					event.srcElement.value = vr.substr( 0, tam - 2  ) + '/' + vr.substr( tam - 2, tam );
				}
				if ( tam >= 5 && tam <= 10 )
				{
					event.srcElement.value = vr.substr( 0, 2 ) + '/' + vr.substr( 2, 2 ) + '/' + vr.substr( 4, 4 ); }
				}
			}
		else
		{
			var s = event.srcElement.value;
			var u=s.length
			u=u-1;
			var ss=s.slice(0,u)
			if(u==0)
			{
				event.srcElement.value = '';
			}
				else
			{
				event.srcElement.value = ss;
			}
		}
	}
}
</script>

<script>
function manda01()
{
window.location="consulta_escopo.asp?selMegaProcesso="+document.frm1.selMegaProcesso.value
}
</script>

<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="foca()">
<form name="frm1" method="post" action="gera_consulta_escopo.asp">
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Enviar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <p align="center"><font color="#000080" face="Verdana" size="3">Consulta de
  Escopo de Cenário</font></p>
  <table border="0" width="973" height="268">
    <tr>
      <td width="149" height="25"> </td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Mega-Processo :</font></b> </td>
      <td width="579" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selMegaProcesso" onChange="manda01()">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL MEGA.EOF=TRUE
          if trim(mega2)=trim(MEGA("MEPR_CD_MEGA_PROCESSO")) then
          %>
          <option selected value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%else%>
          <option value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
          end if
          MEGA.MOVENEXT
          LOOP
          %>
        </select></font></b></td>
    </tr>
    <tr>
      <td width="149" height="21"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione
        o Assunto :&nbsp;</font></b> </td>
      <td width="579" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selAssunto">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL ASSUNTO.EOF=TRUE%>
          <option value="<%=TRIM(ASSUNTO("SUMO_NR_CD_SEQUENCIA"))%>"><%=ASSUNTO("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
			ASSUNTO.MOVENEXT          
          LOOP
          %>
        </select></font></b></td>
    </tr>
    <tr>
      <td width="149" height="21"></td>
      <td width="225" height="21">&nbsp;</td>
      <td width="579" align="left" height="21">&nbsp;</td>
    </tr>
    <tr>
      <td width="149" height="25"> </td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione a Onda :</font></b> </td>
      <td width="579" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selOnda">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL ONDA.EOF=TRUE%>
          <option value="<%=ONDA("ONDA_CD_ONDA")%>"><%=ONDA("ONDA_TX_DESC_ONDA")%></option>
          <%
			ONDA.MOVENEXT          
          LOOP
          %>
        </select></font></b></td>
    </tr>
    <tr>
      <td width="149" height="15"> </td>
      <td width="225" height="15"> </td>
      <td width="579" align="left" height="15"></td>
    </tr>
    <tr>
      <td width="149" height="25"> </td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Status :</font></b> </td>
      <td width="579" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selStatus">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL STATUS.EOF=TRUE%>
          <option value="<%=TRIM(STATUS("SITU_TX_CD_STATUS"))%>"><%=STATUS("SITU_TX_DESC_SITUACAO")%></option>
          <%
			STATUS.MOVENEXT          
          LOOP
          %>
        </select></font></b></td>
    </tr>
    <tr>
      <td width="149" height="21"></td>
      <td width="225" height="21">&nbsp;</td>
      <td width="579" align="left" height="21">&nbsp;</td>
    </tr>
    <tr>
      <td width="149" height="25"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Evento</font></b></td>
      <td width="579" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selEvento" onClick="document.frm1.data01.value=this.value">
          <option value="">== Selecione um Evento ==</option>
          <%
          DO UNTIL EVENTO.EOF=TRUE
          DATA1 = EVENTO("EVEN_DT_EVENTO")
          
          DIA=RIGHT("00"& DAY(DATA1), 2)
          MES=RIGHT("00"& MONTH(DATA1), 2)
          ANO=RIGHT("00"& YEAR(DATA1), 2)
          
          DATA1=DIA & "/" & MES & "/" & "20" & ANO
                    
          %>
          <option value="<%=DATA1%>"><%=DATA1%> - <%=EVENTO("EVEN_TX_DESCRICAO")%></option>
          <%
          EVENTO.MOVENEXT
          LOOP
          %>
        </select></font></b></td>
    </tr>
    <tr>
      <td width="149" height="25"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Ou Digite uma data específica :</font></b></td>
      <td width="579" align="left" height="25">
          <input type="text" name="data01" size="13" maxlength="10" title="Informe a data, formato DD/MM/AAAA: dia com 2 dígitos, mês com 2 dígitos e ano com 4 dígitos">
        <font size="1" face="Verdana, Arial, Helvetica, sans-serif">Formato dd/mm/aaaa 
        - Ex.25/12/2002 - A data m&iacute;nima para este campo &eacute; 06/03/2003</font></td>
    </tr>
  </table>
  <p>&nbsp;</p>
  </form>
</body>

</html>
