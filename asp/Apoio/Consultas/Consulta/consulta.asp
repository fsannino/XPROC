<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../conn_consulta.asp" -->
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

response.write orgao_1_

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
<!-- #include file ="head.asp" -->
<title>....::::::: Sinergia</title>

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
window.location.href="consulta.asp?str01="+document.frm1.Str01.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&modo="+document.frm1.modo.value
}

function manda02()
{
window.location.href="consulta.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&modo="+document.frm1.modo.value
}

function manda03()
{
window.location.href="consulta.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&modo="+document.frm1.modo.value
}

function manda04()
{
window.location.href="consulta.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&str04="+document.frm1.Str04.value+"&selModulo="+document.frm1.selModulo.value+"&op="+document.frm1.op.value+"&atrib="+document.frm1.atrib.value+"&org="+document.frm1.org.value+"&modo="+document.frm1.modo.value
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

<script language="javascript" src="troca_lista.js"></script>

<!-- #include file = "body.asp" -->
<form name="frm1" method="POST" action="gera_consulta.asp">

<table border="0" cellspacing="0" cellpadding="0" width="744">
  <tr> 
    <td height="18" colspan="4" width="460"><img src="_0.gif" width="1" height="1"></td>
  </tr>
  <tr> 
      <td width="1" valign="top"><div align="right"></div></td>
    <td width="76">&nbsp;</td>
    <td width="655"><p><img src="004001.gif" alt=":: Apoiadores Locais"><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif"><br>
          <b>Selecione o Órgão&nbsp;&nbsp;
          </b></font></p>
        <table width="467" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="1">&nbsp;</td>
            <td width="35"><b><font face="Verdana" size="1" color="#000080"><input type="radio" value="1" name="org_" onClick="document.frm1.org.value=this.value" <%=org_1%> checked></font></b></td>
            <td width="180"><b><font face="Verdana" size="1" color="#666666">Órgão Apoiado</font></b></td>
            <td width="7">&nbsp;</td>
            <td width="25">
<b><font face="Verdana" size="1" color="#000080"><input type="radio" value="2" name="org_" onClick="document.frm1.org.value=this.value" <%=org_2%>></font></b>
            </td>
            <td width="210"><b><font face="Verdana" size="1" color="#666666">Órgão 
              de Lota&ccedil;&atilde;o</font></b></td>
          </tr>
        </table>
        <p><b><font face="Georgia, Times New Roman, Times, serif" size="2" color="#666666">Selecione o modo de Consulta</font></b></p>
        <table width="467" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="1">&nbsp;</td>
            <td width="35"><b><font face="Verdana" size="1" color="#000080"><input type="radio" value="1" name="modo_" onClick="document.frm1.modo.value=this.value" <%=modo_1%> checked></font></b></td>
            <td width="180"><b><font face="Verdana" size="1" color="#666666">Apenas 
              o Órgão Selecionado</font></b></td>
            <td width="7">&nbsp;</td>
            <td width="25">
<b><font face="Verdana" size="1" color="#000080"><input type="radio" value="2" name="modo_" onClick="document.frm1.modo.value=this.value" <%=modo_2%>></font></b>
            </td>
            <td width="210"><b><font face="Verdana" size="1" color="#666666">Todos 
              os Órgão Abaixo do Selecionado</font></b></td>
          </tr>
        </table>
      
      <input type="hidden" name="org" size="20" value="<%=org%>">
      <input type="hidden" name="op" size="6" value="<%=request("op")%>">
      <input type="hidden" name="modo" size="20" value="<%=modo%>"><br>
      <input type="hidden" name="selModulo_" size="52">
      
        <table width="467" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="139" height="25"><font color="#666666" face="Verdana" size="1"><b>Atribuição</b></font></td>
      <td width="250" height="25"> 
        <select size="1" name="atrib" style="font-size: 10 px; font-family: Verdana">
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
        </select>
      </td>
      <td width="78" height="25"></td>
    </tr>
    <tr> 
      <td width="139" height="25"><font color="#666666" face="Verdana" size="1"><b>Órgão 
        Aglutinador</b></font></td>
      <td width="250" height="25"> 
        <select size="1" name="Str01" onChange="manda01()" style="font-family: Verdana,Arial; font-size: 10px; bgcolor=" background-color: #6699CC" #ffffff";>
          <OPTION VALUE="0">» Selecione um nível </OPTION>
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
      <td width="78" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="139" height="25"><font color="#666666" face="Verdana" size="1"><b>Órgão 
        de Lotação</b></font></td>
            <td width="250" height="25"> <select size="1" name="Str02" onChange="manda02()" style="font-family: Verdana; font-size: 10 px">
                <option value="000">» Selecione o Nível </option>
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
      <td width="78" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="139" height="25"><font color="#666666" face="Verdana" size="1"><b>Gerência</b></font></td>
      <td width="250" height="25"> 
        <select size="1" name="Str03" onChange="manda03()" style="font-family: Verdana; font-size: 10 px">
          <OPTION VALUE="0">» Selecione o Nível </OPTION>
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
      <td width="78" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="139" height="25"><font color="#666666" face="Verdana" size="1"><b>Gerência 
        Setorial</b></font></td>
      <td width="250" height="25"> 
        <select size="1" name="Str04" onChange="manda04()" style="font-family: Verdana; font-size: 10 px">
          <OPTION VALUE="0">» Selecione o Nível </OPTION>
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
        </select>
      </td>
      <td width="78" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="139" height="25"><font color="#666666" face="Verdana" size="1"><b>Órgão 
        Menor</b></font></td>
      <td width="250" height="25"> 
        <select size="1" name="Str05" style="font-family: Verdana; font-size: 10 px">
          <OPTION VALUE="0">» Selecione o Nível </OPTION>
          <%
        do until str5.eof=true
        %>
          <option value="<%=str5("ORME_CD_ORG_MENOR")%>"><%=str5("ORME_SG_ORG_MENOR")%></option>
          <%
        str5.movenext
        looP
        %>
        </select>
      </td>
      <td width="78" height="25">&nbsp;</td>
    </tr>
  </table>
        <p><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif"><b>Selecione 
          o Assunto</b></font></p>
        <table width="655" border="0" cellspacing="0" cellpadding="0" height="265">
          <tr>
            <td width="78" valign="top" height="116">
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#666666" face="Verdana" size="1"><b>Assuntos 
                Disponíveis&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></p>
            </td>
            <td width="569" colspan="5" height="116"> 
              <select size="6" name="selModulo" multiple style="font-family: Verdana; font-size: 10 px"><%
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
        </select>
            </td>
            <td width="1" valign="top" height="116"><br>
              <p>&nbsp;</p>
            </td>
			<td width="1" height="116">
            </td>
          </tr>
          <tr> 
            <td width="78" valign="top" height="33"></td>
            <td width="188" align="center" height="33"> 
            </td>
            <td width="56" align="center" height="33"><img border="0" src="000029.gif" width="20" height="20" onClick="move(document.frm1.selModulo,document.frm1.list2,1)">
            </td>
            <td width="9" align="center" height="33"> 
            </td>
            <td width="44" align="center" height="33"><img border="0" src="000030.gif" width="20" height="20" onClick="move(document.frm1.list2,document.frm1.selModulo,1)">
            </td>
            <td width="264" align="center" height="33"> 
            </td>
            <td width="1" valign="top" height="33">
            </td>
			<td width="1" height="33">
            </td>
          </tr>
          <tr> 
            <td width="78" valign="top" height="116">
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#666666" face="Verdana" size="1"><b>Assuntos 
                Selecionados &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></td>
            <td width="569" colspan="5" height="116"> 
<select size="6" name="list2" multiple style="font-size: 10 px; font-family: Verdana">
        </select>
            </td>
            <td width="1" valign="top" height="116">
            </td>
			<td width="1" height="116">
            </td>
          </tr>
        </table>
        <table width="75%" border="0" align="center">
          <tr>
    <td width="16%">&nbsp;</td>
            <td width="65%"><div align="center"><strong><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="004203.gif" alt=":: Bot&atilde;o Montar Consulta" onClick="Confirma()"></font></strong></div></td>
    <td width="19%">&nbsp;</td>
  </tr>
</table>

        <p align="center"><strong><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
          </font></strong></p>
        <p align="center"><img src="../000025.gif" width="467" height="1"></p>
      <p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 
        2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p><P>
      </td>
    <td width="4">&nbsp;</td>
  </tr>
</table>

</form>
</html>
