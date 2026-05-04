<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conecta.asp" -->
<%
server.scripttimeout=99999999
response.buffer=false

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

operacao = request("opti")
str_MegaProcesso = request("selMega")

if request("str01")<>"" then
	orgao_1=request("str01")
	orgao_1_= left(formatnumber(ORGAO_1), len(formatnumber(orgao_1))-3)
else
	orgao_1=0
end if

if request("str02")<>"" then
	orgao_2=request("str02")
	orgao_2=right((left(orgao_2,5)),3)	

if(left(orgao_2,1))=0 then
	orgao_2=right(orgao_2,(len(orgao_2))-1)
end if

else
	orgao_2="000"
end if

if request("str03")<>"" then
	orgao_3=request("str03")
else
	orgao_3=0
end if

if request("str04")<>"" then
	orgao_4=request("str04")
else
	orgao_4=0
end if

SSQL1=""
SSQL1="SELECT AGLU_SG_AGLUTINADO, AGLU_CD_AGLUTINADO FROM dbo.ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO"

set str1=db.execute(ssql1)

SSQL2=""
SSQL2="SELECT dbo.ORGAO_MAIOR.ORLO_CD_ORG_LOT, "
SSQL2=SSQL2+"dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT, dbo.ORGAO_MAIOR.ORLO_NR_ORDEM, dbo.ORGAO_MAIOR.ORLO_NM_ORG_LOT, dbo.ORGAO_MAIOR.ORLO_CD_STATUS"
SSQL2=SSQL2+" FROM dbo.ORGAO_MAIOR "
SSQL2=SSQL2+" WHERE (dbo.ORGAO_MAIOR.ORLO_CD_STATUS = 'A') AND (dbo.ORGAO_MAIOR.AGLU_CD_AGLUTINADO = '" & orgao_1_ & "')"
SSQL2=SSQL2+" ORDER BY dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT"
	
set str2=db.execute(ssql2)

ssql3=""
ssql3="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql3=ssql3+" WHERE (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORME_CD_STATUS = 'A')"
ssql3=ssql3+" AND SUBSTRING(ORME_CD_ORG_MENOR,3,3)='" & right("000"& ORGAO_2,3) & "' AND SUBSTRING(ORME_CD_ORG_MENOR,11,5)='00000' AND SUBSTRING(ORME_CD_ORG_MENOR,8,3)<>'000'"

set str3=db.execute(ssql3)

ssql4=""
ssql4="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql4=ssql4+" WHERE (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (ORME_CD_STATUS = 'A')"
ssql4=ssql4+" AND SUBSTRING(ORME_CD_ORG_MENOR,1,10)='" & ORGAO_3 & "' AND SUBSTRING(ORME_CD_ORG_MENOR,14,2)='00' AND SUBSTRING(ORME_CD_ORG_MENOR,13,3) <> '000'" 

set str4=db.execute(ssql4)

ssql5=""
ssql5="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql5=ssql5+" WHERE (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (ORME_CD_STATUS = 'A')"
ssql5=ssql5+" AND SUBSTRING(ORME_CD_ORG_MENOR,1,13)='" & ORGAO_4 & "'  AND SUBSTRING(ORME_CD_ORG_MENOR,14,2)<>'00'" 

set str5=db.execute(ssql5)

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, MEGA_PROCESSO.MEPR_TX_ABREVIA "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " order by MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, MEGA_PROCESSO.MEPR_TX_ABREVIA "

set rsmega=db.execute(str_SQL_MegaProc)

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , MEGA_PROCESSO.MEPR_TX_ABREVIA "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " order by MEGA_PROCESSO.MEPR_TX_ABREVIA "

set rsmega2=db.execute(str_SQL_MegaProc)

SQL_Assunto=""
SQL_Assunto = SQL_Assunto & " SELECT SUMO_NR_CD_SEQUENCIA"
SQL_Assunto = SQL_Assunto & " ,SUMO_TX_DESC_SUB_MODULO"
SQL_Assunto = SQL_Assunto & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
SQL_Assunto = SQL_Assunto & " FROM SUB_MODULO"

if str_MegaProcesso <> 0 then
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaProcesso,2) & "%'" 
else
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS = '9999'"
end if

SQL_Assunto=SQL_Assunto + " ORDER BY SUMO_TX_DESC_SUB_MODULO"

set rsmodulo=db.execute(SQL_Assunto)

%>
<html>
<head>

<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function carrega_txt1(fbox) {
document.frm1.txtfuncao.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtfuncao.value = document.frm1.txtfuncao.value + "'" + fbox.options[i].value + "',";
}
}

function MM_changePropOO(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  var obj2 = MM_findObj(theValue);
  //alert("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
  if (obj && obj2 && (theProp.indexOf("style.")==-1 || obj.style &&  obj2.style )) eval("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->

</script>

<title>#BACKLOG - Solicitações de Melhoria no SAP R/3#</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;

</script>

<script>
function manda()
{
//window.location.href="cad_backlog.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&selMega="+document.frm1.selMega.value+"&selModulo="+document.frm1.selModulo.value+"&selPrioridade="+document.frm1.selPrioridade.value+"&selResponsavel="+document.frm1.selResponsavel.value+"&selTipo="+document.frm1.selTipo.value+"&selLegado="+document.frm1.txtLegado.value
document.frm1.submit()
}

function Confirma() 
{

if((document.frm1.selMega.value==0))
{
alert("Você deve selecionar um MEGA-PROCESSO!");
document.frm1.selMega.focus();
return;
}

if((document.frm1.Str01.value==0)&&(document.frm1.Str02.value==0)&&(document.frm1.Str03.value==0))
{
alert("Você deve Selecionar um ÓRGAO!");
document.frm1.Str01.focus();
return;
}

if((document.frm1.txtTitulo.value==""))
{
alert("Você deve especificar um TÍTULO!");
document.frm1.txtTitulo.focus();
return;
}

if((document.frm1.txtDescricao.value==""))
{
alert("Você deve especificar uma DESCRIÇÃO!");
document.frm1.txtDescricao.focus();
return;
}

if((document.frm1.txtSolicitante.value==""))
{
alert("Você deve especificar o NOME do SOLICITANTE!");
document.frm1.txtSolicitante.focus();
return;
}

if((document.frm1.txtChave.value==""))
{
alert("Você deve especificar a CHAVE do SOLICITANTE!");
document.frm1.txtChave.focus();
return;
}

if((document.frm1.txtFone.value==""))
{
alert("Você deve especificar o TELEFONE do SOLICITANTE!");
document.frm1.txtFone.focus();
return;
}

if((document.frm1.selResponsavel.value==0))
{
alert("Você deve selecionar um RESPONSAVEL NO SINERGIA!");
document.frm1.selResponsavel.focus();
return;
}

if((document.frm1.selPrioridade.value==0))
{
alert("Você deve selecionar uma PRIORIDADE!");
document.frm1.selPrioridade.focus();
return;
}

if((document.frm1.selTipo.value==0))
{
alert("Você deve selecionar um TIPO!");
document.frm1.selTipo.focus();
return;
}

else
{
document.frm1.action='valida_cad_backlog.asp';
document.frm1.submit();
}
}

function verifica_tecla(e)
{
if(window.event.keyCode==16)
{
alert("Tecla não permitida!");
return;
}
}

function Verifica_tamanho(e)
{
var tam = e.length;
if(tam>1000)
{
	tam=1000;
	var txt = e.slice(0,1000);
	document.frm1.txtDescricao.value = txt;
}
document.frm1.txttamanho.value=1000-tam;
}

</SCRIPT>


<script language="javascript" src="troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF">
<form name="frm1" method="POST" action="cad_backlog.asp">

  <table border="0" width="96%" height="244">
    <tr> 
      <td width="7%" height="33">&nbsp;</td>
      <td width="18%" height="33">&nbsp;</td>
      <td height="33" colspan="2"> 
        <p align="left"><font face="Verdana" color="#000080">Solicita&ccedil;&otilde;es 
          de Melhoria na Solu&ccedil;&atilde;o Configurada no SAP R/3</font> 
      </td>
    </tr>
    <tr> 
      <td width="7%" height="23">&nbsp;</td>
      <td width="18%" height="23">&nbsp;</td>
      <td height="23" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td width="7%" height="1"></td>
      <td width="18%" height="1"></td>
      <td height="1" colspan="2"></td>
    </tr>
    <tr> 
      <td width="7%" height="18"></td>
      <td width="18%" height="18"><b><font face="System" size="2" color="#000080">Mega-Processo</font></b></td>
      <td height="18" colspan="2"> 
        <select name="selMega" style="font-family: Verdana; font-size: 8 pt" onChange="manda()">
          <OPTION VALUE="0">== Selecione o Mega-Processo ==</OPTION>
          <%
		  do until rsmega.eof=true
		  if trim(str_MegaProcesso)= trim(rsmega("MEPR_CD_MEGA_PROCESSO")) then
		  	selec="Selected"
		  else
		  	selec=""
		  end if
		  %>
          <OPTION <%=selec%> VALUE=<%=rsmega("MEPR_CD_MEGA_PROCESSO")%>><%=rsmega("MEPR_TX_DESC_MEGA_PROCESSO")%></OPTION>
          <%
		  rsmega.movenext
		  loop
		  %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="20"></td>
      <td width="18%" height="20"><b><font face="System" size="2" color="#000080">Assunto</font></b></td>
      <td height="20" colspan="2"> 
        <select name="selModulo" style="font-family: Verdana; font-size: 8 pt">
          <option value="0">== Selecione o Assunto ==</option>
          <%
		  do until rsmodulo.eof=true
		  if trim(request("selModulo"))=trim(rsmodulo("SUMO_NR_CD_SEQUENCIA")) then
			selecm="selected"		  
		  else
		  	selecm=""
		  end if
		  %>
          <OPTION <%=selecm%> VALUE="<%=rsmodulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rsmodulo("SUMO_TX_DESC_SUB_MODULO")%></OPTION>
          <%
		  rsmodulo.movenext
		  loop
		  %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="18"></td>
      <td width="18%" height="18">&nbsp;</td>
      <td height="18" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td width="7%" height="18"></td>
      <td width="18%" height="18"><font color="#000080" face="System" size="2"><b>Órgão</b></font></td>
      <td height="18" colspan="2"> 
        <select size="1" name="Str01" onChange="manda()" style="font-family: Verdana; font-size: 8 pt">
          <OPTION VALUE="0">== Selecione Órgão Aglutinador ==</OPTION>
          <%do until str1.eof=true
        if trim(orgao_1)=trim(str1("AGLU_CD_AGLUTINADO")) then
        %>
          <option selected value="<%=str1("AGLU_CD_AGLUTINADO")%>"><%=str1("AGLU_SG_AGLUTINADO")%></option>
          <%else%>
          <option value="<%=str1("AGLU_CD_AGLUTINADO")%>"><%=str1("AGLU_SG_AGLUTINADO")%></option>
          <%
        end if
        str1.movenext
        looP
        %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="24"></td>
      <td width="18%" height="24"><b><font face="System" size="2" color="#000080">Unidade</font></b></td>
      <td height="24" colspan="2"> 
        <select size="1" name="Str02" onChange="manda()" style="font-family: Verdana; font-size: 8 pt">
          <OPTION VALUE="000">== Selecione Órgão de Lotação ==</OPTION>
          <%do until str2.eof=true
        if trim(orgao_2)=trim(str2("ORLO_CD_ORG_LOT"))then
        %>
          <option selected value="<%=ORGAO_1 & right(("000" & str2("ORLO_CD_ORG_LOT")),3) & right(("000" & str2("ORLO_NR_ORDEM")),2)%>"><%=str2("ORLO_SG_ORG_LOT")%></option>
          <%else%>
          <option value="<%=ORGAO_1 & right(("000" & str2("ORLO_CD_ORG_LOT")),3) & right(("000" & str2("ORLO_NR_ORDEM")),2)%>"><%=str2("ORLO_SG_ORG_LOT")%></option>
          <%
        end if
        str2.movenext
        looP
        %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="13"></td>
      <td width="18%" height="13"><font color="#000080" face="System" size="2"><b> 
        Gerência</b></font></td>
      <td height="13" colspan="2"> 
        <select size="1" name="Str03" style="font-family: Verdana; font-size: 8 pt">
          <OPTION VALUE="0">== Selecione Gerência ==</OPTION>
          <%
        do until str3.eof=true
        if trim(orgao_3)=trim(left((str3("ORME_CD_ORG_MENOR")),10)) then
        %>
          <option selected value="<%=trim(left((str3("ORME_CD_ORG_MENOR")),10))%>"><%=str3("ORME_SG_ORG_MENOR")%></option>
          <%else%>
          <option value="<%=trim(left((str3("ORME_CD_ORG_MENOR")),10))%>"><%=str3("ORME_SG_ORG_MENOR")%></option>
          <%

        end if
        str3.movenext
        looP
        %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="23">&nbsp;</td>
      <td width="18%" height="23" valign="top">&nbsp;</td>
      <td height="23" valign="bottom" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td width="7%" height="42">&nbsp;</td>
      <td width="18%" height="42" valign="middle"><b><font face="System" size="2" color="#000080">T&iacute;tulo 
        da Solicita&ccedil;&atilde;o</font></b></td>
      <td height="42" valign="middle" colspan="2"> 
        <input type="text" name="txtTitulo" size="100" maxlength="100" style="font-family: Verdana; font-size: 8 pt" value="<%=request("txtTitulo")%>">
      </td>
    </tr>
    <tr> 
      <td width="7%" height="72"></td>
      <td width="18%" height="72" valign="top"><b><font face="System" size="2" color="#000080">Descri&ccedil;&atilde;o 
        da Solicita&ccedil;&atilde;o</font></b></td>
      <td height="72" colspan="2"> 
        <textarea name="txtDescricao" cols="100" rows="7" onKeyUp="Verifica_tamanho(this.value)"><%=request("txtDescricao")%></textarea>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="30"></td>
      <td width="18%" height="30">&nbsp;</td>
      <td height="30" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><font color="#000099">Tamanho 
        M&aacute;ximo : 1000 caracteres - Restantes :</font></b></font> 
        <input type="text" name="txttamanho" size="7" maxlength="4" disabled>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="40"></td>
      <td width="18%" height="40"><b><font face="System" size="2" color="#000080"> 
        Solicitante</font></b></td>
      <td height="40" width="8%"><b><font face="System" size="2" color="#000080">Nome</font></b> 
      </td>
      <td height="40" width="67%"> 
        <input type="text" name="txtSolicitante" size="70" maxlength="50" style="font-family: Verdana; font-size: 8 pt" value="<%=request("txtSolicitante")%>">
      </td>
    </tr>
    <tr> 
      <td width="7%" height="31">&nbsp;</td>
      <td width="18%" height="31">&nbsp;</td>
      <td height="31" width="8%"><b><font face="System" size="2" color="#000080">Chave</font></b> 
      </td>
      <td height="31" width="67%"> 
        <input type="text" name="txtChave" size="10" maxlength="4" style="font-family: Verdana; font-size: 8 pt" value="<%=request("txtChave")%>">
      </td>
    </tr>
    <tr> 
      <td width="7%" height="31">&nbsp;</td>
      <td width="18%" height="31">&nbsp;</td>
      <td height="31" width="8%"><b><font face="System" size="2" color="#000080">Telefone</font></b> 
      </td>
      <td height="31" width="67%"> 
        <input type="text" name="txtFone" size="15" maxlength="9" style="font-family: Verdana; font-size: 8 pt" value="<%=request("txtFone")%>">
      </td>
    </tr>
    <tr> 
      <td width="7%" height="20">&nbsp;</td>
      <td width="18%" height="20">&nbsp;</td>
      <td height="20" colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td width="7%" height="31">&nbsp;</td>
      <td width="18%" height="31"><b><font face="System" size="2" color="#000080">Respons&aacute;vel 
        no Sinergia</font></b></td>
      <td height="31" colspan="2"> 
        <select size="1" name="selResponsavel" style="font-family: Verdana; font-size: 8 pt">
          <option value="0">== Selecione o Responsável ==</option>
          <%
		  do until rsmega2.eof=true
		  if trim(request("selResponsavel"))= trim(rsmega2("MEPR_CD_MEGA_PROCESSO")) then
		  	selec="Selected"
		  else
		  	selec=""
		  end if
		  %>
          <OPTION <%=selec%> VALUE=<%=rsmega2("MEPR_CD_MEGA_PROCESSO")%>><%=rsmega2("MEPR_TX_ABREVIA")%></OPTION>
          <%
		  rsmega2.movenext
		  loop
		  %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="22">&nbsp;</td>
      <td width="18%" height="22">&nbsp;</td>
      <td height="22" colspan="2">&nbsp; </td>
    </tr>
    <tr> 
      <td width="7%" height="28">&nbsp;</td>
      <td width="18%" height="28"><b><font size="2" color="#000080"><font face="System">Prioridade</font></font></b></td>
      <td height="28" colspan="2"> 
        <select size="1" name="selPrioridade" style="font-family: Verdana; font-size: 8 pt">
          <%
			select case request("selPrioridade")
			case 1
				chp0=""			
				chp1="selected"
				chp2=""								
			case 2
				chp0=""			
				chp1=""
				chp2="selected"								
			case else
				chp0="selected"			
				chp1=""
				chp2=""								
			end select
		  %>
          <option <%=chp0%> value="0">== Selecione a Prioridade ==</option>
          <option <%=chp1%> value="1">Pr&eacute; GoLive Sinergia</option>
          <option <%=chp2%> value="2">P&oacute;s GoLive</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="9">&nbsp;</td>
      <td width="18%" height="9"><b><font size="2" color="#000080"><font face="System">Tipo</font></font></b> 
      </td>
      <td height="9" colspan="2"> 
        <select size="1" name="selTipo" style="font-family: Verdana; font-size: 8 pt">
          <%
			select case request("selTipo")
			case 1
				cht0=""			
				cht1="selected"
				cht2=""								
				cht3=""								
			case 2
				cht0=""			
				cht1=""
				cht2="selected"								
				cht3=""								
			case 3
				cht0=""			
				cht1=""
				cht2=""								
				cht3="selected"								
			case else
				cht0="selected"			
				cht1=""
				cht2=""								
				cht3=""								
			end select
		  %>
          <option <%=cht0%> value="0">== Selecione o Tipo ==</option>
          <option <%=cht1%> value="1">Corretiva</option>
          <option <%=cht2%> value="2">Legal / Normativa</option>
          <option <%=cht3%> value="3">Melhoria</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="7%" height="9">&nbsp;</td>
      <td width="18%" height="9">&nbsp;</td>
      <td width="8%" height="9">&nbsp;</td>
      <td width="67%" height="9">&nbsp;</td>
    </tr>
    <tr> 
      <td width="7%" height="9">&nbsp;</td>
      <td width="18%" height="9"><b><font size="2" color="#000080"><font face="System">Existe 
        no Legado ?</font></font></b></td>
      <td width="8%" height="9"> 
        <%
	 if request("selLegado")=2 then
	  	checa1=""
		checa2="checked"
	else
	  	checa2=""
		checa1="checked"
	end if
	  %>
        <div align="center"><b> 
          <input type="radio" name="selLegado" value="1" <%=checa1%> onClick="document.frm1.txtLegado.value=1">
          </b></div>
      </td>
      <td width="67%" height="9"><b><font size="2" color="#000080"><font face="System">Sim</font></font></b></td>
    </tr>
    <tr> 
      <td width="7%" height="9">&nbsp;</td>
      <td width="18%" height="9">&nbsp;</td>
      <td width="8%" height="9"> 
        <div align="center"><b> 
          <input type="radio" name="selLegado" value="2" <%=checa2%> onClick="document.frm1.txtLegado.value=2">
          </b></div>
      </td>
      <td width="67%" height="9"><b><font size="2" color="#000080"><font face="System">N&atilde;o</font></font></b></td>
    </tr>
    <tr> 
      <td width="7%" height="9">&nbsp;</td>
      <td width="18%" height="9">&nbsp;</td>
      <td height="9" colspan="2"> 
        <%
		if request("selLegado")="" then
			Legado=1
		else
			Legado = request("selLegado")
		end if
	  %>
        <input type="hidden" name="txtLegado" value="<%=Legado%>">
      </td>
    </tr>
    <tr> 
      <td width="7%" height="28">&nbsp;</td>
      <td width="18%" height="28"><b><font size="2" color="#000080"></font></b></td>
      <td height="28" colspan="2"> 
        <input type="button" name="Submit" value="Gravar Solicitação" onClick="Confirma()">
      </td>
    </tr>
  </table>
  <p align="left">
</form>
<%
db.close
set db = nothing
%>
</body>
<script>
{
Verifica_tamanho(document.frm1.txtDescricao.value)
}
</script>  
</html>