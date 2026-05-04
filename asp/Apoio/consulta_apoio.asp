<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn_consulta.asp" -->
<%
operacao = request("opti")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

visual=request("visual")

select case visual
case 1
	op1="selected"
	op2=""
	op3=""
case 2
	op1=""
	op2="selected"
	op3=""
case 3
	op1=""
	op2=""
	op3="selected"
end select


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

ssql="SELECT DISTINCT "
ssql=ssql+"(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 2)) AS FAIXA0, "
ssql=ssql+"(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 8),2)) AS FAIXA1 , "
ssql=ssql+"(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 5),2)) AS FAIXA2 , "
ssql=ssql+"(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 2)) AS FAIXA3, "
ssql=ssql+"dbo.MEGA_PROCESSO.MEPR_TX_ABREVIA AS FAIXA4 , dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, "
ssql=ssql+"dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA,  dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
ssql=ssql+"FROM dbo.APOIO_LOCAL_MODULO INNER JOIN dbo.SUB_MODULO ON "
ssql=ssql+"dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
ssql=ssql+"LEFT JOIN dbo.MEGA_PROCESSO ON "
ssql=ssql+"(CONVERT(char(3),dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO)) IN ((REPLACE((REPLACE(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS,'0','')),'-',', '))) "
ssql=ssql+"ORDER BY FAIXA0, FAIXA1, FAIXA2, FAIXA3, FAIXA4, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "

set rs_modulo=db.execute(ssql)

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
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
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
window.location.href="consulta_apoio.asp?str01="+document.frm1.Str01.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&visual="+document.frm1.visual.value
}

function manda02()
{
window.location.href="consulta_apoio.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&visual="+document.frm1.visual.value
}

function manda03()
{
window.location.href="consulta_apoio.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&visual="+document.frm1.visual.value
}

function manda04()
{
window.location.href="consulta_apoio.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&str04="+document.frm1.Str04.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&visual="+document.frm1.visual.value
}

function Confirma()
{
if((document.frm1.Str01.value==0)&&(document.frm1.Str02.value==0)&&(document.frm1.Str03.value==0))
{
alert('Você deve selecionar um Órgão!');
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
</script><script language="JavaScript">

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
</script><script language="javascript" src="troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onKeyDown="verifica_tecla()">

<form name="frm1" method="POST" action="gera_consulta_apoiador.asp">
   <input type="hidden" name="org" size="20" value="<%=org%>"><input type="hidden" name="op" size="6" value="<%=request("op")%>"><p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; </p>
   <table border="0" width="79%">
              <tr>
                         <td width="71%"><p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2" color="#000080">Consulta de Órgãos Apoiados x Assunto</font></b></td>
              </tr>
   </table>
   <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; </p>
   <table border="0" width="614" height="1">
              <tr>
                         <td width="33" height="1"></td>
                         <td width="186" height="1"><font color="#000080" face="Verdana" size="1"><b>Modo de Visualização</b></font></td>
                         <td width="359" height="1"><select size="1" name="visual" style="font-family: Verdana; font-size: 8 pt">
                            <option <%=op1%> value="1">QUANTITATIVO</option>
                            <option <%=op2%> value="2">NOMINAL DETALHADO</option>
                            <option <%=op3%> value="3">NOMINAL ACUMULATIVO</option>
                            </select></td>
                         <td width="79" height="1"></td>
                         <td width="59" height="1"></td>
                         <td width="50" height="19">&nbsp;</td>
              </tr>
              <tr>
                         <td width="33" height="1"></td>
                         <td width="186" height="1"><font color="#000080" face="Verdana" size="1"><b>Atribuição</b></font></td>
                         <td width="359" height="1"><select size="1" name="atrib" style="font-family: Verdana; font-size: 8 pt"><%if request("atrib")=1 then%>
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
				          end if%></select></td>
                         <td width="79" height="1"></td>
                         <td width="59" height="1"></td>
                         <td width="50" height="19">&nbsp;</td>
              </tr>
              <tr>
                         <td width="33" height="1"></td>
                         <td width="186" height="1"></td>
                         <td width="359" height="1"></td>
                         <td width="79" height="1"></td>
                         <td width="59" height="1"></td>
                         <td width="50" height="19">&nbsp;</td>
              </tr>
              <tr>
                         <td width="33" height="1"></td>
                         <td width="186" height="1"><font color="#000080" face="Verdana" size="1"><b>Órgão Aglutinador</b></font></td>
                         <td width="359" height="1"><select size="1" name="Str01" style="font-family: Verdana; font-size: 8 pt" onChange="manda01()">
                            <option VALUE="0">== Selecione um nível ==</option>
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
        %></select></td>
                         <%if request("op")=0 then%> <td width="79" height="1">&nbsp;</td>
                         <%else%> <td width="59" height="1">&nbsp;</td>
                         <%end if%> <td width="50" rowspan="3" height="44"></td>
              </tr>
              <tr>
                         <td width="33" height="1"></td>
                         <td width="186" height="1"><font color="#000080" face="Verdana" size="1"><b>Órgão de Lotação</b></font></td>
                         <td width="359" height="1"><select size="1" name="Str02" style="font-family: Verdana; font-size: 8 pt" onChange="manda02()">
                            <option VALUE="000">== Selecione o Nível ==</option>
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
        %></select></td>
              </tr>
              <tr>
                         <td width="33" height="1"></td>
                         <td width="186" height="1"><font color="#000080" face="Verdana" size="1"><b>Gerência</b></font></td>
                         <td width="359" height="1"><select size="1" style="font-family: Verdana; font-size: 8 pt" name="Str03">
                            <option VALUE="0">== Selecione o Nível ==</option>
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
        %></select></td>
                         <td width="79" height="17" align="center"><b></b></td>
              </tr>
   </table>
   <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
   <table border="0" width="79%" height="156">
              <tr>
                         <td width="2%" height="15"><p style="margin-top: 0; margin-bottom: 0"></td>
                         <td width="98%" colspan="4" height="15"><p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" size="2" color="#5F76AD">Selecione o Assunto referente ao&nbsp; Apoiador Local/Multiplicador</font></b></p>
                            <p style="margin-top: 0; margin-bottom: 0"></p>
                            <p style="margin-top: 0; margin-bottom: 0"></td>
              </tr>
              <tr>
                         <td width="2%" height="21"></td>
                         <td width="2%" height="21">
                            <div align="right">
&nbsp;</div>
                         </td>
                         <td width="41%" height="21"><font color="#000080" face="Verdana" size="1"><b>Assuntos disponíveis</b></font><font size="1"> <input type="hidden" name="selModulo_" size="52"></font></td>
                         <td width="9%" height="21"></td>
                         <td width="47%" height="21">
                            <div align="left">
                                      <font color="#000080" face="Verdana" size="1"><b>Assuntos selecionados</b></font></div>
                         </td>
              </tr>
              <tr>
                         <td width="2%" height="108" rowspan="5" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
                         <td width="2%" height="108" rowspan="5" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#000080" face="Verdana" size="2">&nbsp;</font></td>
                         <td width="41%" height="108" rowspan="5" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font size="1" face="Latha">
                         <select size="6" name="selModulo" multiple style="font-family: Verdana; font-size: 8 pt">
                         <option value="0">== TODOS OS ASSUNTOS ==</option>
                            <%
        DO UNTIL RS_MODULO.EOF=TRUE
        IF RS_MODULO("SUMO_NR_CD_SEQUENCIA")<>33 AND RS_MODULO("SUMO_NR_CD_SEQUENCIA")<>34 AND RS_MODULO("SUMO_NR_CD_SEQUENCIA")<>36 THEN
        
		IF NOT ISNULL(RS_MODULO("FAIXA4")) THEN
		%>
                            <option value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("FAIXA4")%>-<%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
                            <%END IF%> <%
		IF (RS_MODULO("FAIXA0")=RS_MODULO("FAIXA1")) AND (RS_MODULO("FAIXA0")=RS_MODULO("FAIXA2")) AND (RS_MODULO("FAIXA0")<>RS_MODULO("FAIXA3")) THEN
		%>
                            <option value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("FAIXA0")%>-<%=RS_MODULO("FAIXA3")%>-<%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
                            <%END IF%> <%
		IF (RS_MODULO("FAIXA0")=RS_MODULO("FAIXA1")) AND (RS_MODULO("FAIXA0")<>RS_MODULO("FAIXA2")) AND (RS_MODULO("FAIXA0")<>RS_MODULO("FAIXA3")) THEN
		%>
                            <option value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("FAIXA0")%>-<%=RS_MODULO("FAIXA2")%>-<%=RS_MODULO("FAIXA3")%>-<%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
                            <%END IF%> <%
		IF (RS_MODULO("FAIXA0")<>RS_MODULO("FAIXA1")) AND (RS_MODULO("FAIXA0")<>RS_MODULO("FAIXA2")) AND (RS_MODULO("FAIXA0")<>RS_MODULO("FAIXA3")) THEN
		%>
                            <option value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("FAIXA0")%>-<%=RS_MODULO("FAIXA1")%>-<%=RS_MODULO("FAIXA2")%>-<%=RS_MODULO("FAIXA3")%>-<%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
                            <%END IF%> <%
        END IF
        RS_MODULO.MOVENEXT
        LOOP
        %></select></font></td>
                         <td width="9%" height="24" align="center" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><a href="#" onClick="move(document.frm1.selModulo,document.frm1.list2,1)"><img border="0" src="../../imagens/continua_F01.gif" width="24" height="24"></a></td>
                         <td width="47%" height="108" rowspan="5" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><select size="6" name="list2" multiple style="font-size: 8 pt; font-family: Verdana"></select></td>
              </tr>
              <tr>
                         <td width="9%" height="23" align="center" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
              </tr>
              <tr>
                         <td width="9%" height="23" align="center" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><a href="#" onClick="move(document.frm1.list2,document.frm1.selModulo,1)">
                            <img border="0" src="../../imagens/continua2_F01.gif" width="24" height="24"></a></td>
              </tr>
              <tr>
                         <td width="9%" height="23" align="center" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
              </tr>
              <tr>
                         <td width="9%" height="15" align="center" valign="top"><p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
              </tr>
   </table>
   <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><input type="button" value="Montar Consulta" name="B1" onClick="Confirma()"> </p>
</form>
</body>

</html>