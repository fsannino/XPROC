<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="conn_consulta.asp" -->
<%
operacao = request("opti")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("org")=2 then
	org=2
	org_1=""
	org_2="checked"
else
	org=1
	org_1="checked"
	org_2=""
end if

if request("modo")=2 then
	modo=2
	modo_1=""
	modo_2="checked"
else
	modo=1
	modo_1="checked"
	modo_2=""
end if

set rs_modulo=db.execute("SELECT * FROM " & Session("Prefixo") & "SUB_MODULO ORDER BY SUMO_TX_ABREV, SUMO_TX_DESC_SUB_MODULO")

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

SET str1=db.execute(ssql1)

SSQL2=""
SSQL2="SELECT dbo.ORGAO_MAIOR.ORLO_CD_ORG_LOT, "
SSQL2=SSQL2+"dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT, dbo.ORGAO_MAIOR.ORLO_NR_ORDEM, dbo.ORGAO_MAIOR.ORLO_NM_ORG_LOT, dbo.ORGAO_MAIOR.ORLO_CD_STATUS"
SSQL2=SSQL2+" FROM dbo.ORGAO_MAIOR "
SSQL2=SSQL2+" WHERE (dbo.ORGAO_MAIOR.ORLO_CD_STATUS = 'A') AND (dbo.ORGAO_MAIOR.AGLU_CD_AGLUTINADO = '" & orgao_1_ & "')"
SSQL2=SSQL2+" ORDER BY dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT"

set str2=db.execute(ssql2)

ssql3=""
ssql3="SELECT ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql3=ssql3+" WHERE (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORME_CD_STATUS = 'A')"
ssql3=ssql3+" AND SUBSTRING(ORME_CD_ORG_MENOR,3,3)='" & right("000"& ORGAO_2,3) & "' AND SUBSTRING(ORME_CD_ORG_MENOR,11,5)='00000' AND SUBSTRING(ORME_CD_ORG_MENOR,8,3)<>'000' ORDER BY ORME_SG_ORG_MENOR"

set str3=db.execute(ssql3)

ssql4=""
ssql4="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql4=ssql4+" WHERE (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (ORME_CD_STATUS = 'A')"
ssql4=ssql4+" AND SUBSTRING(ORME_CD_ORG_MENOR,1,10)='" & ORGAO_3 & "' AND SUBSTRING(ORME_CD_ORG_MENOR,14,2)='00' AND SUBSTRING(ORME_CD_ORG_MENOR,13,3) <> '000'  ORDER BY ORME_SG_ORG_MENOR" 

set str4=db.execute(ssql4)

ssql5=""
ssql5="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql5=ssql5+" WHERE (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (ORME_CD_STATUS = 'A')"
ssql5=ssql5+" AND SUBSTRING(ORME_CD_ORG_MENOR,1,13)='" & ORGAO_4 & "'  AND SUBSTRING(ORME_CD_ORG_MENOR,14,2)<>'00' ORDER BY ORME_SG_ORG_MENOR" 

set str5=db.execute(ssql5)
%>
<html>
<head>

<title>Base de Dados de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<script>
function carrega_txt(fbox) 
{
document.frm1.selModulo_.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.selModulo_.value = document.frm1.selModulo_.value + "," + fbox.options[i].value;
}
}

function manda01()
{
window.location.href="consulta_apoio.asp?str01="+document.frm1.Str01.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&modo="+document.frm1.modo.value+"&momento="+document.frm1.momento.value
}

function manda02()
{
window.location.href="consulta_apoio.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&modo="+document.frm1.modo.value+"&momento="+document.frm1.momento.value
}

function manda03()
{
window.location.href="consulta_apoio.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&modo="+document.frm1.modo.value+"&momento="+document.frm1.momento.value
}

function manda04()
{
window.location.href="consulta_apoio.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&str04="+document.frm1.Str04.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&modo="+document.frm1.modo.value+"&momento="+document.frm1.momento.value
}

function Confirma()
{
if((document.frm1.Str01.value==0)&&(document.frm1.Str02.value==0)&&(document.frm1.Str03.value==0)&&(document.frm1.Str04.value==0)&&(document.frm1.Str05.value==0)&&(document.frm1.list2.options.length == 0))
{
alert('Pelo menos um parâmetro de consulta deve ser especificado!');
return;
}
else
{
carrega_txt(document.frm1.list2);
var i = document.frm1.selModulo_.value;
var t = i.length
document.frm1.selModulo_.value = i.slice(1,t)
document.frm1.submit()
}
}
</SCRIPT>

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

function verifica_tecla(e)
{
if(window.event.keyCode==16)
{
alert("Tecla não permitida!");
return;
}
if(window.event.keyCode==121)
{
window.history.go(-1);
return;
}
if(window.event.keyCode==120)
{
window.print();
}
}
</script>

<script language="javascript" src="troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onKeyDown="verifica_tecla()">
<form name="frm1" method="POST" action="gera_consulta_apoiador.asp">

  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;
   </p>
  <table border="0" width="100%">
    <tr>
      <td width="71%">
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font size="5" color="#000080" face="Tahoma"><b>Consulta
  Apoiador Local / Multiplicador</b></font>
      </td>
      <td width="29%">
        <%if request("op")=1 then%>
        <p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Arial Narrow" size="3"><a href="menu.asp"><img border="0" src="../../imagens/volta_f02.gif" align="absmiddle"></a></font><font size="2"><font face="Arial Narrow">
        </font><b><font color="#000080" face="Verdana">Menu Principal</font></b></font></td>
        <%end if%>
    </tr>
  </table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;
   </p>
  <table border="0" width="100%" height="49">
    <tr>
      <td width="10%" height="19"></td>
      <td width="60%" colspan="4" height="19"><b><font face="Verdana" size="2" color="#5F76AD">Selecione
        o Órgão&nbsp;</font></b></td>
      <td width="30%" height="19"></td>
    </tr>
    <tr>
      <td width="10%" height="18"><input type="hidden" name="op" size="6" value="<%=request("op")%>"></td>
      <td width="4%" height="18">
        <p align="center"><b><font face="Verdana" size="1" color="#000080"><input type="radio" value="1" name="org_" onClick="document.frm1.org.value=this.value" <%=org_1%>></font></b></td>
      <td width="13%" height="18"><b><font face="Verdana" size="1" color="#000080">&nbsp;Órgão
        Apoiado</font></b></td>
      <td width="3%" height="18">
        <p align="center"><b><font face="Verdana" size="1" color="#000080"><input type="radio" value="2" name="org_" onClick="document.frm1.org.value=this.value" <%=org_2%>></font></b></td>
      <td width="40%" height="18"><b><font face="Verdana" size="1" color="#000080">&nbsp;Órgão
        de Lotação</font></b></td>
      <td width="30%" height="18">
      <input type="hidden" name="org" size="20" value="<%=org%>"></td>
    </tr>
  </table>
  <table border="0" width="100%" height="49">
    <tr>
      <td width="10%" height="19"></td>
      <td width="71%" colspan="4" height="19"><b><font face="Verdana" size="2" color="#5F76AD">Selecione
        o modo de Consulta&nbsp;</font></b></td>
      <td width="19%" height="19"></td>
    </tr>
    <tr>
      <td width="10%" height="18"></td>
      <td width="4%" height="18">
        <p align="center"><b><font face="Verdana" size="1" color="#000080"><input type="radio" value="1" name="modo_" onClick="document.frm1.modo.value=this.value" <%=modo_1%>></font></b></td>
      <td width="22%" height="18"><b><font face="Verdana" size="1" color="#000080">&nbsp;Apenas
        o Órgão Selecionado</font></b></td>
      <td width="3%" height="18">
        <p align="center"><b><font face="Verdana" size="1" color="#000080"><input type="radio" value="2" name="modo_" onClick="document.frm1.modo.value=this.value" <%=modo_2%>></font></b></td>
      <td width="42%" height="18"><b><font face="Verdana" size="1" color="#000080">&nbsp;Todos
        os Órgão Abaixo do Selecionado</font></b></td>
      <td width="19%" height="18">
      <input type="hidden" name="modo" size="20" value="<%=modo%>"></td>
    </tr>
  </table>
  <table border="0" width="638" height="62">
    <tr>
      <td width="71" height="25"></td>
      <td width="138" height="25"><font color="#000080" face="Verdana" size="2"><b>Atribuição</b></font></td>
      <td width="322" height="25"><select size="1" name="atrib">
          <%if request("atrib")=1 then%>
          <option selected value="1">APOIADOR LOCAL</option>
          <option value="2">MULTIPLICADOR</option>
          <%else
          if request("atrib")=2 then
          %>
          <option value="1">APOIADOR LOCAL</option>
          <option selected value="2">MULTIPLICADOR</option>
          <%
          else
          %>
          <option selected value="1">APOIADOR LOCAL</option>
          <option value="2">MULTIPLICADOR</option>
          <%
          end if
          end if
          %>
        </select></td>
      <td width="82" height="25"></td>
    </tr>
    <tr>
      <td width="71" height="29"></td>
      <td width="138" height="29"><font color="#000080" face="Verdana" size="2"><b>Momento</b></font></td>
      <td width="322" height="29"><select size="1" name="momento">
          <%if request("momento")=0 then%>
          <option selected value="0">TODOS</option>
          <option value="1">MOMENTO 1</option>
          <option value="2">MOMENTO 2</option>
          <option value="12">MOMENTO 1 e 2</option>
          <%
          else
          if request("momento")=1 then
          %>
          <option value="0">TODOS</option>
          <option selected value="1">MOMENTO 1</option>
          <option value="2">MOMENTO 2</option>
          <option value="12">MOMENTO 1 e 2</option>
          <%else%>
          <%if request("momento")=2 then%>
          <option value="0">TODOS</option>
          <option value="1">MOMENTO 1</option>
          <option selected value="2">MOMENTO 2</option>
          <option value="12">MOMENTO 1 e 2</option>
          <%else%>
          <%if request("momento")=12 then%>
          <option value="0">TODOS</option>
          <option value="1">MOMENTO 1</option>
          <option value="2">MOMENTO 2</option>
          <option selected value="12">MOMENTO 1 e 2</option>
          <%else%>
          <option value="0">TODOS</option>
          <option value="1">MOMENTO 1</option>
          <option value="2">MOMENTO 2</option>
          <option value="12">MOMENTO 1 e 2</option>
          <%
          end if
          end if
          end if
          end if
          %>
        </select></td>
      <td width="82" height="29"></td>
    </tr>
  </table>
  <table border="0" width="783" height="138">
    <tr> 
      <td width="72" height="1"></td>
      <td width="139" height="1"><font color="#000080" face="Verdana" size="2"><b>Órgão 
        Aglutinador</b></font></td>
      <td width="339" height="1"><select size="1" name="Str01" onChange="manda01()">
          <OPTION VALUE="0">== Selecione um nível ==</OPTION>
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
        </select></td>
      <%if request("op")=0 then%>
      <td width="174" height="1">&nbsp;</td>
      <%else%>
      <td width="174" height="1">&nbsp;</td>
      <%end if%>
      <td width="37" rowspan="4" height="113"> </td>
    </tr>
    <tr> 
      <td width="72" height="18"></td>
      <td width="139" height="18"><font color="#000080" face="Verdana" size="2"><b>Órgão 
        de Lotação</b></font></td>
      <td width="339" height="18"><select size="1" name="Str02" onChange="manda02()">
          <OPTION VALUE="000">== Selecione o Nível ==</OPTION>
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
        </select></td>
    </tr>
    <tr> 
      <td width="72" height="25"></td>
      <td width="139" height="25"><font color="#000080" face="Verdana" size="2"><b> 
        Gerência</b></font></td>
      <td width="339" height="25"><select size="1" name="Str03" onChange="manda03()">
          <OPTION VALUE="0">== Selecione o Nível ==</OPTION>
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
        </select></td>
      <td width="174" height="50" rowspan="2" align="center"><b></b></td>
    </tr>
    <tr> 
      <td width="72" height="25"></td>
      <td width="139" height="25"><font color="#000080" face="Verdana" size="2"><b> 
        Gerência Setorial</b></font></td>
      <td width="339" height="25"><select size="1" name="Str04" onChange="manda04()">
          <OPTION VALUE="0">== Selecione o Nível ==</OPTION>
          <%
        do until str4.eof=true
        if trim(orgao_4)=trim(left((str4("ORME_CD_ORG_MENOR")),13)) then
        %>
          <option selected value="<%=LEFT((str4("ORME_CD_ORG_MENOR")),13)%>"><%=str4("ORME_SG_ORG_MENOR")%></option>
          <%else%>
          <option value="<%=LEFT((str4("ORME_CD_ORG_MENOR")),13)%>"><%=str4("ORME_SG_ORG_MENOR")%></option>
          <%

        end if
        str4.movenext
        looP
        %>
        </select></td>
    </tr>
    <tr> 
      <td width="72" height="25"></td>
      <td width="139" height="25"><font color="#000080" face="Verdana" size="2"><b>Órgão 
        Menor</b></font></td>
      <td width="339" height="25"><select size="1" name="Str05">
          <OPTION VALUE="0">== Selecione o Nível ==</OPTION>
          <%
        do until str5.eof=true
        %>
          <option value="<%=str5("ORME_CD_ORG_MENOR")%>"><%=str5("ORME_SG_ORG_MENOR")%></option>
          <%
        str5.movenext
        looP
        %>
        </select></td>
      <td width="174" height="25"></td>
    </tr>
    <tr> 
      <td width="72" height="25"></td>
      <td width="139" height="25"></td>
      <td height="25" colspan="2"></td>
    </tr>
  </table>
  <table border="0" width="94%" height="156">
    <tr>
      <td width="12%" height="15"></td>
      <td width="48%" colspan="2" height="15"><b><font face="Verdana" size="2" color="#5F76AD">Selecione 
        o Assunto referente ao&nbsp; Apoiador Local/Multiplicador</font></b></td>
      <td width="5%" height="15"></td>
      <td width="35%" height="15"></td>
    </tr>
    <tr>
      <td width="12%" height="21"></td>
      <td width="13%" height="21"><div align="right"><font color="#000080" face="Verdana" size="2"></font></div></td>
      <td width="36%" height="21"><font color="#000080" face="Verdana" size="2"><b>Assuntos 
        dispon&iacute;veis</b></font> 
        <input type="hidden" name="selModulo_" size="52"></td>
      <td width="5%" height="21"></td>
      <td width="35%" height="21"><div align="left"><font color="#000080" face="Verdana" size="2"><b>Assuntos 
          selecionados</b></font></div></td>
    </tr>
    <tr>
      <td width="12%" height="108" rowspan="5" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
      <td width="13%" height="108" rowspan="5" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#000080" face="Verdana" size="2">&nbsp;</font></p>
      </td>
      <td width="36%" height="108" rowspan="5" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><select size="6" name="selModulo" multiple><%
        DO UNTIL RS_MODULO.EOF=TRUE
        IF RS_MODULO("SUMO_NR_CD_SEQUENCIA")<>33 AND RS_MODULO("SUMO_NR_CD_SEQUENCIA")<>34 AND RS_MODULO("SUMO_NR_CD_SEQUENCIA")<>36 THEN
        if trim(Modulo)=trim(RS_MODULO("SUMO_NR_CD_SEQUENCIA"))then
        %>
        <option selected value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("SUMO_TX_ABREV")%>-<%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
        <%else%>
        <option value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("SUMO_TX_ABREV")%>-<%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
        <%
        end if
        END IF
        RS_MODULO.MOVENEXT
        LOOP
        %>
        </select></p>
      </td>
      <td width="5%" height="24" align="center" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><a href="#" onClick="move(document.frm1.selModulo,document.frm1.list2,1)"><img border="0" src="../../imagens/continua_F01.gif" width="24" height="24"></a></p>
      </td>
      <td width="35%" height="108" rowspan="5" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><select size="6" name="list2" multiple>
        </select></p>
      </td>
    </tr>
    <tr>
      <td width="5%" height="23" align="center" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    </tr>
    <tr>
      <td width="5%" height="23" align="center" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><a href="#" onClick="move(document.frm1.list2,document.frm1.selModulo,1)"><img border="0" src="../../imagens/continua2_F01.gif" width="24" height="24"></a></p>
      </td>
    </tr>
    <tr>
      <td width="5%" height="23" align="center" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    </tr>
    <tr>
      <td width="5%" height="15" align="center" valign="top">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    </tr>
  </table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
<input type="button" value="Montar Consulta" name="B1" onClick="Confirma()">
</form>
</body>
</html>
