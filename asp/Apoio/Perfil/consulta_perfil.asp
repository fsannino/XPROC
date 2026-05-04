<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn_consulta.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

if request("caso") = 1 then
	pre_funcao = "HR.%"
else
	pre_funcao = "MM.%"
end if

if request("exibe")= 1  then
	habilita = "disabled"
end if

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


str01 = request("str01")
str02 = request("str02")
str03 = request("str03")

orgao=""
tem_o = 0
	
if str01<>0 then
	orgao = str01
end if

if str02<>"000" then
	orgao = str02
end if

if str03<>0 then
	orgao = str03
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

SET str2=db.execute(ssql2)

ssql3=""
ssql3="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql3=ssql3+" WHERE (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORME_CD_STATUS = 'A')"
ssql3=ssql3+" AND SUBSTRING(ORME_CD_ORG_MENOR,3,3)='" & right("000"& ORGAO_2,3) & "' AND SUBSTRING(ORME_CD_ORG_MENOR,11,5)='00000' AND SUBSTRING(ORME_CD_ORG_MENOR,8,3)<>'000'"

set str3=db.execute(ssql3)

if request("exibe")=1 then
if pre_funcao = "HR.%" then
	ssql=""
	ssql="SELECT DISTINCT "
	ssql=ssql+"dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO, dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	ssql=ssql+"FROM dbo.USMA_MICRO_R3_VISAO_R3 "
	ssql=ssql+"INNER JOIN dbo.USUARIO_MAPEAMENTO ON "
	ssql=ssql+"dbo.USMA_MICRO_R3_VISAO_R3.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
	ssql=ssql+"INNER JOIN dbo.ORGAO_MENOR ON "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR "
	ssql=ssql+"INNER JOIN dbo.MACRO_PERFIL ON "
	ssql=ssql+"dbo.USMA_MICRO_R3_VISAO_R3.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
	ssql=ssql+"INNER JOIN dbo.FUNCAO_NEGOCIO ON "
	ssql=ssql+"dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
	ssql=ssql+"WHERE dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & orgao & "%' AND dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO LIKE '" & pre_funcao & "'"
	ssql=ssql+"ORDER BY dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO, dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO	"
else
	ssql=""
	ssql="SELECT DISTINCT "
	ssql=ssql+"dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO, dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	ssql=ssql+"FROM dbo.USUARIO_PERFIL "
	ssql=ssql+"INNER JOIN dbo.USUARIO_MAPEAMENTO ON "
	ssql=ssql+"dbo.USUARIO_PERFIL.USPE_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
	ssql=ssql+"INNER JOIN dbo.ORGAO_MENOR ON "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR "
	ssql=ssql+"INNER JOIN dbo.MACRO_PERFIL ON "
	ssql=ssql+"dbo.USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
	ssql=ssql+"INNER JOIN dbo.FUNCAO_NEGOCIO ON "
	ssql=ssql+"dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
	ssql=ssql+"WHERE dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & orgao & "%' AND dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO LIKE '" & pre_funcao & "'"
	ssql=ssql+"ORDER BY dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO, dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO	"
end if
else
	ssql="select distinct fune_cd_funcao_negocio from macro_perfil where fune_cd_funcao_negocio like 'aa.%'"
end if

'response.write ssql

set funcao = db.execute(ssql)

reg_func = funcao.RecordCount
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Usuários cadastrados com Perfil</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
<script>
function manda01()
{
window.location.href="consulta_perfil.asp?str01="+document.frm1.Str01.value+"&Caso="+document.frm1.caso.value+"&exibe=0"
}

function manda02()
{
window.location.href="consulta_perfil.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&Caso="+document.frm1.caso.value+"&exibe=0"
}

function manda03()
{
window.location.href="consulta_perfil.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&Caso="+document.frm1.caso.value+"&exibe=0"
}

function manda04()
{
if((document.frm1.Str01.value==0)&&(document.frm1.Str02.value==0)&&(document.frm1.Str03.value==0))
{
alert('Você deve selecionar um Parâmetro de Consulta!');
document.frm1.Str01.focus();
return;
}
else
{
window.location.href="consulta_perfil.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&exibe=1&Caso="+document.frm1.caso.value
}
}

function Confirma()
{
if((document.frm1.Str01.value==0)&&(document.frm1.Str02.value==0)&&(document.frm1.Str03.value==0))
{
alert('Você deve selecionar um Parâmetro de Consulta!');
document.frm1.Str01.focus();
return;
}
else
{
document.frm1.target='_top';
document.frm1.submit()
}
}
</script>

<script language="javascript" src="troca_lista.js"></script>

<body bgcolor="#E5E5E5" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<form name="frm1" method="POST" action="gera_consulta_perfil.asp">
   <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; </p>
   <table border="0" width="75%">
              <tr><td width="71%"><p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#000080">Usuários cadastrados no R/3 com Perfil</font></b></td>
              </tr>
		   </table>
			   <table border="0" width="729" height="97" style="border-collapse: collapse" bordercolor="#111111" cellpadding="2" cellspacing="3">
		              <tr>
                         <td width="33" height="25" bgcolor="#E5E5E5" bordercolor="#E5E5E5"></td>
                         <td width="189" height="25">
                         <input type="hidden" name="visual" size="4" value="3">
                         <input type="hidden" name="atrib" size="3" value="1">
                         </td>
                         <td width="531" height="25">
                         
                         <input type="hidden" name="caso" size="20" value="<%=request("caso")%>">
                         
                         </td>
                         <td width="17" height="25"></td>
		              </tr>
        		      <tr>
                         <td width="33" height="20" bgcolor="#E5E5E5" bordercolor="#E5E5E5"></td>
                         <td width="189" height="20"><font color="#000080" face="Verdana" size="2"><b>Órgão Aglutinador</b></font></td>
                         <td width="531" height="20"><select size="1" name="Str01" style="font-family: Verdana; font-size: 7 pt" onChange="manda01()">
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
        					%>
        					</select></td>
                          </tr>
			              <tr>
                         <td width="33" height="29" bgcolor="#E5E5E5" bordercolor="#E5E5E5"></td>
                         <td width="189" height="29"><font color="#000080" face="Verdana" size="2"><b>Órgão</b></font></td>
                         <td width="531" height="29"><select size="1" name="Str02" style="font-family: Verdana; font-size: 7 pt" onChange="manda02()">
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
					        looP%>
					        </select></td>
			              </tr>
			              <tr>
                         <td width="33" height="20" bgcolor="#E5E5E5" bordercolor="#E5E5E5">&nbsp;</td>
                         <td width="189" height="20"><b><font face="Verdana" size="2" color="#000080">Órgão Menor</font></b></td>
                         <td width="531" height="20"><select size="1" name="Str03" style="font-family: Verdana; font-size: 7 pt" onChange="manda03()">
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
              		</tr>
			        <%if request("exibe")=1 then%>
			        <tr>
                         <td width="33" height="26" bgcolor="#E5E5E5" bordercolor="#E5E5E5">&nbsp;</td>
                         <td width="189" height="26"><b><font face="Verdana" size="2" color="#000080">Função</font></b></td>
                         <td width="531" height="26"><select size="1" name="selFuncao" style="font-family: Verdana; font-size: 7 pt">
                            <option value="N">== Selecione a Função ==</option>
                            <%
                            do until funcao.eof=true
                            if trim(request("Funcao")) = trim(funcao("fune_cd_funcao_negocio")) then
                            	checa = "selected"
                            else
                            	checa=""
                            end if
                            
                            TITULO = funcao("fune_tx_titulo_funcao_negocio")
                            
                            IF RIGHT(TRIM(TITULO),6) = "ANTIGA" THEN
                            	TITULO = LEFT(TITULO,LEN(TITULO)-7)
                            	TITULO = TITULO & "ANTECIP"
                            ELSE
                            	TITULO=TITULO
                            END IF
                            %>
                            <option <%=checa%> value="<%=funcao("fune_cd_funcao_negocio")%>"><%=funcao("fune_cd_funcao_negocio")%> - <%=LEFT(TITULO,60)%></option>
                            <%
                            funcao.movenext
                            loop
                            %>
                            </select></td>
              		</tr>
          		   	<%end if%>
          		   	</table>
&nbsp;
<%
if request("exibe")=1 then
if reg_func > 0  then
%>
	<p align="center"><input type="button" value="Montar Consulta" name="B2" onClick="Confirma()"></p>
<%
	else
%>
	&nbsp;<p align="center"><b><font size="2" face="Verdana" color="#800000">Nenhuma Função Encontrada para a Seleção!
<%
	end if
else
%> </font></b></p>
<p align="center"><input type="button" value="Carregar Funções" name="B1" onClick="manda04()"></p>
<%
end if
%>
</form>
</body>

</html>