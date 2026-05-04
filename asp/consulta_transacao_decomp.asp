<<<<<<< HEAD
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set mega=db.execute("SELECT * FROM MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set evento=db.execute("SELECT * FROM EVENTO ORDER BY EVEN_DT_EVENTO")

if request("selMegaProcesso")<>0 then
	strmega=request("selMegaProcesso")
else
	strmega=0
end if

if request("selProcesso")<>0 then
	strproc=request("selProcesso")
else
	strproc=0
end if

if request("selSubProcesso")<>0 then
	strsub=request("selSubProcesso")
else
	strsub=0
end if

if request("selAtividade")<>0 then
	strativ=request("selAtividade")
else
	strativ=0
end if

set processo=db.execute("SELECT * FROM PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & strmega & " ORDER BY PROC_TX_DESC_PROCESSO")

set subprocesso=db.execute("SELECT * FROM SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & strmega & " AND PROC_CD_PROCESSO=" & strproc & " ORDER BY SUPR_TX_DESC_SUB_PROCESSO")

SSQL=""
SSQL="SELECT DISTINCT * FROM dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, dbo.RELACAO_FINAL.PROC_CD_PROCESSO, dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, dbo.ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA, dbo.ATIVIDADE_CARGA.ATCA_TX_DESC_ATIVIDADE FROM dbo.RELACAO_FINAL INNER JOIN"
SSQL=SSQL+" dbo.ATIVIDADE_CARGA ON dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = dbo.ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA "
SSQL=SSQL+"WHERE (dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strmega & ") AND (dbo.RELACAO_FINAL.PROC_CD_PROCESSO = " & strproc & ") AND (dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & strsub & ") ORDER BY dbo.ATIVIDADE_CARGA.ATCA_TX_DESC_ATIVIDADE"

set atividade=db.execute(SSQL)

ssql=""
ssql="SELECT DISTINCT * FROM dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, dbo.RELACAO_FINAL.PROC_CD_PROCESSO, "
ssql=ssql+"dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, dbo.MODULO_R3.MODU_CD_MODULO, "
ssql=ssql+"dbo.MODULO_R3.MODU_TX_DESC_MODULO FROM dbo.RELACAO_FINAL INNER JOIN dbo.MODULO_R3 ON dbo.RELACAO_FINAL.MODU_CD_MODULO = dbo.MODULO_R3.MODU_CD_MODULO "
ssql=ssql+"WHERE (dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strmega & ") AND (dbo.RELACAO_FINAL.PROC_CD_PROCESSO = " & strproc & ") "
ssql=ssql+"AND (dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & strsub & ") AND (dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & strativ & ")"

set modulo=db.execute(SSQL)

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
<%if request("tipo")=1 then%>
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
alert('Data de Referência Inválida! Tente novamente');
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
<%else%>
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

var a=document.frm1.data02.value
var chk2   = 0;
var maxDay = 0;

var dd = a.slice(0,2)
var mm = a.slice(3,5)
var yyyy = a.slice(6,10)

maxDay = max_day(mm, yyyy);  

if((dd <= 0) || (dd > maxDay))
{ chk2 = 1;}
else if((mm <= 0) || (mm > 12))
{ chk2 = 1;}
else if((yyyy <= 0))
{ chk2 = 1;} 


if(chk == 1)
{ 
alert('Data Inicial Inválida! Tente novamente');
document.frm1.data01.value='';
document.frm1.data01.focus();
return;
}
if(chk2 == 1)
{ 
alert('Data Final Inválida! Tente novamente');
document.frm1.data02.value='';
document.frm1.data02.focus();
return;
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
<%end if%>

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
window.location="consulta_transacao_decomp.asp?tipo="+document.frm1.tipo.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value
}

function manda02()
{
window.location="consulta_transacao_decomp.asp?tipo="+document.frm1.tipo.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value
}

function manda03()
{
window.location="consulta_transacao_decomp.asp?tipo="+document.frm1.tipo.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selsubProcesso="+document.frm1.selSubProcesso.value
}

function manda04()
{
window.location="consulta_transacao_decomp.asp?tipo="+document.frm1.tipo.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selsubProcesso="+document.frm1.selSubProcesso.value+"&selAtividade="+document.frm1.selAtividade.value
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
<form name="frm1" method="post" action="gera_consulta_transacao_decomp.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
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
  Escopo de Transações</font></p>
  <table border="0" width="822">
    <tr>
      <td width="130"> <input type="hidden" name="tipo" size="10" value="<%=request("tipo")%>"> </td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione o Mega-Processo :</font></b> </td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selMegaProcesso" onChange="manda01()">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL MEGA.EOF=TRUE
          if trim(request("selMegaProcesso"))=trim(MEGA("MEPR_CD_MEGA_PROCESSO")) then
          %>
          <option selected value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
          else
          %>
          <option value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
          end if
          MEGA.MOVENEXT
          LOOP
          %>
        </select></font></b></td>
    </tr>
    <tr>
      <td width="130"> </td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione
        o Processo:</font></b> </td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selProcesso" onChange="manda02()">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL PROCESSO.EOF=TRUE
          if trim(request("selProcesso"))=trim(processo("PROC_CD_PROCESSO")) then
          %>
          <option selected value="<%=processo("PROC_CD_PROCESSO")%>"><%=processo("PROC_TX_DESC_PROCESSO")%></option>
          <%
          else
          %>
          <option value="<%=processo("PROC_CD_PROCESSO")%>"><%=processo("PROC_TX_DESC_PROCESSO")%></option>
          <%
          end if
          processo.MOVENEXT
          LOOP
          %>
          </select></font></b></td>
    </tr>
    <tr>
      <td width="130"> </td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione o
        Sub-Processo :</font></b> </td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selSubProcesso" onChange="manda03()">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL SUBPROCESSO.EOF=TRUE
          if trim(request("selsubProcesso"))=trim(subprocesso("SUPR_CD_SUB_PROCESSO")) then
          %>
          <option selected value="<%=subprocesso("SUPR_CD_SUB_PROCESSO")%>"><%=subprocesso("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%
          else
          %>
          <option value="<%=subprocesso("SUPR_CD_SUB_PROCESSO")%>"><%=subprocesso("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%
          end if
          SUBPROCESSO.MOVENEXT
          LOOP
          %>
          </select></font></b></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione
        a Atividade :</font></b></td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selAtividade" onChange="manda04()">
          <option value="0">== TODOS ==</option>
          <%do until atividade.eof=true
          if trim(request("selAtividade"))=trim(atividade("ATCA_CD_ATIVIDADE_CARGA")) then
          %>
          <option selected value="<%=atividade("ATCA_CD_ATIVIDADE_CARGA")%>"><%=atividade("ATCA_TX_DESC_ATIVIDADE")%></option>
          <%else%>
          <option value="<%=atividade("ATCA_CD_ATIVIDADE_CARGA")%>"><%=atividade("ATCA_TX_DESC_ATIVIDADE")%></option>
          <%
          end if
          atividade.movenext
          loop
          %>
          </select></font></b></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione o
        Módulo :</font></b></td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selModulo">
          <option value="0">== TODOS ==</option>
          <%do until modulo.eof=true%>
          <option value="<%=modulo("MODU_CD_MODULO")%>"><%=modulo("MODU_TX_DESC_MODULO")%></option>
          <%
          modulo.movenext
          loop
          %>

          </select></font></b></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="244">&nbsp;</td>
      <td width="430" align="left">&nbsp;</td>
    </tr>
    <%if request("tipo")=1 then%>
   <tr>
      <td width="130"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Evento
        ou</font></b></td>
      <td width="430" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selEvento" onClick="document.frm1.data01.value=this.value">
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
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">digite
        uma Data de
        Referência : </font></b></td>
      <td width="430" align="left">
          <input type="text" name="data01" size="15" maxlength="10" title="Informe a data, formato DD/MM/AAAA: dia com 2 dígitos, mês com 2 dígitos e ano com 4 dígitos">
          <font color="#000080" face="Verdana" size="1">Formato DD/MM/AAAA</font></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="244"></td>
      <td width="430" align="left">
      </td>
    </tr>
	<%else%> 
    <tr>
      <td width="130"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Evento
        ou</font></b></td>
      <td width="430" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selEvento" onClick="document.frm1.data01.value=this.value">
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
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">digite a Data
        Inicial : </font></b></td>
      <td width="430" align="left">
          <input type="text" name="data01" size="15" maxlength="10" title="Informe a data, formato DD/MM/AAAA: dia com 2 dígitos, mês com 2 dígitos e ano com 4 dígitos">
          <font color="#000080" face="Verdana" size="1">Formato DD/MM/AAAA</font></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Evento
        ou</font></b></td>
      <td width="430" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selEvento" onClick="document.frm1.data02.value=this.value">
          <option value="">== Selecione um Evento ==</option>
          <%
          evento.movefirst
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
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">digite a Data
        Final : </font></b></td>
      <td width="430" align="left">
          <input type="text" name="data02" size="15" maxlength="10" title="Informe a data, formato DD/MM/AAAA: dia com 2 dígitos, mês com 2 dígitos e ano com 4 dígitos">
          <font color="#000080" face="Verdana" size="1">Formato DD/MM/AAAA</font></td>
    </tr>
   <%end if%>
  </table>
  <p>&nbsp;</p>
  </form>
</body>

</html>
=======
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set mega=db.execute("SELECT * FROM MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set evento=db.execute("SELECT * FROM EVENTO ORDER BY EVEN_DT_EVENTO")

if request("selMegaProcesso")<>0 then
	strmega=request("selMegaProcesso")
else
	strmega=0
end if

if request("selProcesso")<>0 then
	strproc=request("selProcesso")
else
	strproc=0
end if

if request("selSubProcesso")<>0 then
	strsub=request("selSubProcesso")
else
	strsub=0
end if

if request("selAtividade")<>0 then
	strativ=request("selAtividade")
else
	strativ=0
end if

set processo=db.execute("SELECT * FROM PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & strmega & " ORDER BY PROC_TX_DESC_PROCESSO")

set subprocesso=db.execute("SELECT * FROM SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & strmega & " AND PROC_CD_PROCESSO=" & strproc & " ORDER BY SUPR_TX_DESC_SUB_PROCESSO")

SSQL=""
SSQL="SELECT DISTINCT * FROM dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, dbo.RELACAO_FINAL.PROC_CD_PROCESSO, dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, dbo.ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA, dbo.ATIVIDADE_CARGA.ATCA_TX_DESC_ATIVIDADE FROM dbo.RELACAO_FINAL INNER JOIN"
SSQL=SSQL+" dbo.ATIVIDADE_CARGA ON dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = dbo.ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA "
SSQL=SSQL+"WHERE (dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strmega & ") AND (dbo.RELACAO_FINAL.PROC_CD_PROCESSO = " & strproc & ") AND (dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & strsub & ") ORDER BY dbo.ATIVIDADE_CARGA.ATCA_TX_DESC_ATIVIDADE"

set atividade=db.execute(SSQL)

ssql=""
ssql="SELECT DISTINCT * FROM dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, dbo.RELACAO_FINAL.PROC_CD_PROCESSO, "
ssql=ssql+"dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, dbo.MODULO_R3.MODU_CD_MODULO, "
ssql=ssql+"dbo.MODULO_R3.MODU_TX_DESC_MODULO FROM dbo.RELACAO_FINAL INNER JOIN dbo.MODULO_R3 ON dbo.RELACAO_FINAL.MODU_CD_MODULO = dbo.MODULO_R3.MODU_CD_MODULO "
ssql=ssql+"WHERE (dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strmega & ") AND (dbo.RELACAO_FINAL.PROC_CD_PROCESSO = " & strproc & ") "
ssql=ssql+"AND (dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & strsub & ") AND (dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & strativ & ")"

set modulo=db.execute(SSQL)

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
<%if request("tipo")=1 then%>
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
alert('Data de Referência Inválida! Tente novamente');
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
<%else%>
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

var a=document.frm1.data02.value
var chk2   = 0;
var maxDay = 0;

var dd = a.slice(0,2)
var mm = a.slice(3,5)
var yyyy = a.slice(6,10)

maxDay = max_day(mm, yyyy);  

if((dd <= 0) || (dd > maxDay))
{ chk2 = 1;}
else if((mm <= 0) || (mm > 12))
{ chk2 = 1;}
else if((yyyy <= 0))
{ chk2 = 1;} 


if(chk == 1)
{ 
alert('Data Inicial Inválida! Tente novamente');
document.frm1.data01.value='';
document.frm1.data01.focus();
return;
}
if(chk2 == 1)
{ 
alert('Data Final Inválida! Tente novamente');
document.frm1.data02.value='';
document.frm1.data02.focus();
return;
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
<%end if%>

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
window.location="consulta_transacao_decomp.asp?tipo="+document.frm1.tipo.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value
}

function manda02()
{
window.location="consulta_transacao_decomp.asp?tipo="+document.frm1.tipo.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value
}

function manda03()
{
window.location="consulta_transacao_decomp.asp?tipo="+document.frm1.tipo.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selsubProcesso="+document.frm1.selSubProcesso.value
}

function manda04()
{
window.location="consulta_transacao_decomp.asp?tipo="+document.frm1.tipo.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selsubProcesso="+document.frm1.selSubProcesso.value+"&selAtividade="+document.frm1.selAtividade.value
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
<form name="frm1" method="post" action="gera_consulta_transacao_decomp.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
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
  Escopo de Transações</font></p>
  <table border="0" width="822">
    <tr>
      <td width="130"> <input type="hidden" name="tipo" size="10" value="<%=request("tipo")%>"> </td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione o Mega-Processo :</font></b> </td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selMegaProcesso" onChange="manda01()">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL MEGA.EOF=TRUE
          if trim(request("selMegaProcesso"))=trim(MEGA("MEPR_CD_MEGA_PROCESSO")) then
          %>
          <option selected value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
          else
          %>
          <option value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
          end if
          MEGA.MOVENEXT
          LOOP
          %>
        </select></font></b></td>
    </tr>
    <tr>
      <td width="130"> </td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione
        o Processo:</font></b> </td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selProcesso" onChange="manda02()">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL PROCESSO.EOF=TRUE
          if trim(request("selProcesso"))=trim(processo("PROC_CD_PROCESSO")) then
          %>
          <option selected value="<%=processo("PROC_CD_PROCESSO")%>"><%=processo("PROC_TX_DESC_PROCESSO")%></option>
          <%
          else
          %>
          <option value="<%=processo("PROC_CD_PROCESSO")%>"><%=processo("PROC_TX_DESC_PROCESSO")%></option>
          <%
          end if
          processo.MOVENEXT
          LOOP
          %>
          </select></font></b></td>
    </tr>
    <tr>
      <td width="130"> </td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione o
        Sub-Processo :</font></b> </td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selSubProcesso" onChange="manda03()">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL SUBPROCESSO.EOF=TRUE
          if trim(request("selsubProcesso"))=trim(subprocesso("SUPR_CD_SUB_PROCESSO")) then
          %>
          <option selected value="<%=subprocesso("SUPR_CD_SUB_PROCESSO")%>"><%=subprocesso("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%
          else
          %>
          <option value="<%=subprocesso("SUPR_CD_SUB_PROCESSO")%>"><%=subprocesso("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%
          end if
          SUBPROCESSO.MOVENEXT
          LOOP
          %>
          </select></font></b></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione
        a Atividade :</font></b></td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selAtividade" onChange="manda04()">
          <option value="0">== TODOS ==</option>
          <%do until atividade.eof=true
          if trim(request("selAtividade"))=trim(atividade("ATCA_CD_ATIVIDADE_CARGA")) then
          %>
          <option selected value="<%=atividade("ATCA_CD_ATIVIDADE_CARGA")%>"><%=atividade("ATCA_TX_DESC_ATIVIDADE")%></option>
          <%else%>
          <option value="<%=atividade("ATCA_CD_ATIVIDADE_CARGA")%>"><%=atividade("ATCA_TX_DESC_ATIVIDADE")%></option>
          <%
          end if
          atividade.movenext
          loop
          %>
          </select></font></b></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">Selecione o
        Módulo :</font></b></td>
      <td width="430" align="left"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selModulo">
          <option value="0">== TODOS ==</option>
          <%do until modulo.eof=true%>
          <option value="<%=modulo("MODU_CD_MODULO")%>"><%=modulo("MODU_TX_DESC_MODULO")%></option>
          <%
          modulo.movenext
          loop
          %>

          </select></font></b></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="244">&nbsp;</td>
      <td width="430" align="left">&nbsp;</td>
    </tr>
    <%if request("tipo")=1 then%>
   <tr>
      <td width="130"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Evento
        ou</font></b></td>
      <td width="430" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selEvento" onClick="document.frm1.data01.value=this.value">
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
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">digite
        uma Data de
        Referência : </font></b></td>
      <td width="430" align="left">
          <input type="text" name="data01" size="15" maxlength="10" title="Informe a data, formato DD/MM/AAAA: dia com 2 dígitos, mês com 2 dígitos e ano com 4 dígitos">
          <font color="#000080" face="Verdana" size="1">Formato DD/MM/AAAA</font></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="244"></td>
      <td width="430" align="left">
      </td>
    </tr>
	<%else%> 
    <tr>
      <td width="130"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Evento
        ou</font></b></td>
      <td width="430" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selEvento" onClick="document.frm1.data01.value=this.value">
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
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">digite a Data
        Inicial : </font></b></td>
      <td width="430" align="left">
          <input type="text" name="data01" size="15" maxlength="10" title="Informe a data, formato DD/MM/AAAA: dia com 2 dígitos, mês com 2 dígitos e ano com 4 dígitos">
          <font color="#000080" face="Verdana" size="1">Formato DD/MM/AAAA</font></td>
    </tr>
    <tr>
      <td width="130"></td>
      <td width="225" height="25"><b><font color="#000080" face="Verdana" size="2">Selecione o Evento
        ou</font></b></td>
      <td width="430" align="left" height="25"><b><font color="#000080" face="Verdana" size="2"><select size="1" name="selEvento" onClick="document.frm1.data02.value=this.value">
          <option value="">== Selecione um Evento ==</option>
          <%
          evento.movefirst
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
      <td width="130"></td>
      <td width="244"><b><font color="#000080" face="Verdana" size="2">digite a Data
        Final : </font></b></td>
      <td width="430" align="left">
          <input type="text" name="data02" size="15" maxlength="10" title="Informe a data, formato DD/MM/AAAA: dia com 2 dígitos, mês com 2 dígitos e ano com 4 dígitos">
          <font color="#000080" face="Verdana" size="1">Formato DD/MM/AAAA</font></td>
    </tr>
   <%end if%>
  </table>
  <p>&nbsp;</p>
  </form>
</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
