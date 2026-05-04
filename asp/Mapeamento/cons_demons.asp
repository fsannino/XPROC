<!--#include file="conecta.asp" -->
<%
set objUSR = server.createobject("Seseg.Usuario")

if objUSR.GetUsuario then
	chave=objUSR.sei_chave
	lotacao=objUSR.sei_lotacao
	nome=objUSR.sei_nome
	set objUSR = nothing
else
	response.redirect "erro.asp?op=3"
end if

set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")

set db2 = server.createobject("ADODB.CONNECTION")
db2.open "Provider=Microsoft.Jet.Oledb.4.0;data source=" & server.mappath("base.mdb")

tipo_cons="CONSULTA DEMONSTRATIVA POR ÓRGÃO"
legenda = "SELECIONE O ÓRGÃO AGLUTINADOR DESEJADO"
if request("selOrgao")<>"87" then
	legenda2 = "SELECIONE O CURSO DESEJADO"
else
	legenda2 = "SELECIONE O MEGA-PROCESSO DESEJADO"
end if
pagina = "gera_cons_demons.asp"

set rs = db.execute("SELECT AGLU_CD_AGLUTINADO, AGLU_SG_AGLUTINADO FROM ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")

if request("selOrgao")="" then
	set rs1 = db2.execute("SELECT DISTINCT CURSO FROM AB WHERE CURSO='XXXX' ORDER BY CURSO")
else
	if request("selOrgao")="87" then
	
	set rs1 = db.execute("SELECT DISTINCT MEPR_CD_MEGA_PROCESSO, MEPR_TX_DESC_MEGA_PROCESSO FROM MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")	
	
	else
	
	set temp = db.execute("SELECT DISTINCT AGLU_SG_AGLUTINADO AS ORME_SG_ORG_MENOR FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & request("selOrgao"))
	TABELA = "[" & temp("ORME_SG_ORG_MENOR") &"]"
	
	on error resume next
	set rs1 = db2.execute("SELECT DISTINCT CURSO FROM " & TABELA & " ORDER BY CURSO")
	
	if rs1.eof=true or err.number<>0 then
		set rs1 = db2.execute("SELECT DISTINCT CURSO FROM AB WHERE CURSO='XXXX' ORDER BY CURSO")
	end if
	
	end if
	
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<script>
function envia()
{
window.location = 'cons_demons.asp?selOrgao='+document.frm1.selOrgao.value
}
</script>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body topmargin="0" leftmargin="0" link="#000080" vlink="#000080" alink="#000080">
<form name="frm1" method="post" action="<%=pagina%>">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="89%" id="AutoNumber2" height="514">
           <tr>
                      <td width="100%" height="68" valign="top" colspan="2"><img border="0" src="topo.jpg"></td>
           </tr>
           <tr>
                      <td width="13%" height="445" valign="top"><img border="0" src="lado.jpg" width="83" height="429"></td>
                      <td width="87%" height="445" valign="top">
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="93%" id="AutoNumber3" height="65">
                         <tr>
                                    <td width="19%" height="130" align="center" colspan="2"><img border="0" src="mult_c.jpg" align="right"></td>
                                    <td width="81%" height="130" align="left"><font face="Verdana" color="#800000"><b><%=tipo_cons%></b></font></td>
                         </tr>
                         <tr>
                                    <td width="11%" height="29" align="center">&nbsp;</td>
                                    <td width="8%" height="29" align="center"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
                                    <td width="81%" height="29" align="left"><b><font face="Verdana" size="2"><%=legenda%></font></b></td>
                         </tr>
                         <tr>
                                    <td width="11%" height="24" align="center">&nbsp;</td>
                                    <td width="8%" height="24" align="center">&nbsp;</td>
                                    <td width="81%" height="24" align="left">
                                    <select size="1" name="selOrgao" style="font-family: Verdana; font-size: 7 pt" onChange="envia()">
                                       <%
                                       do until rs.eof=true
                                       
                                       if trim(request("selOrgao"))=trim(rs.fields(0).value) then
                                       	ver="selected"                                       
                                       else
                                       	ver=""
                                       end if
                                       
                                       %>
                                       <option <%=ver%> value="<%=rs.fields(0).value%>"><%=rs.fields(1).value%></option>   
                                       <%                                    
                                       rs.movenext
                                       loop
                                       %>
                                       </select></td>
                         </tr>
                         <tr>
                                    <td width="11%" height="24" align="center">&nbsp;</td>
                                    <td width="8%" height="24" align="center">&nbsp;</td>
                                    <td width="81%" height="29" align="left"><b><font face="Verdana" size="2"><%=legenda2%></font></b></td>
                         </tr>
                         <tr>
                                    <td width="11%" height="24" align="center">&nbsp;</td>
                                    <td width="8%" height="24" align="center">&nbsp;</td>
                                    <td width="81%" height="24" align="left">
                                    <select size="1" name="selCurso" style="font-family: Verdana; font-size: 7 pt">
                                    <option value="XXXX">== TODOS ==</option>
                                       <%
                                       if request("selOrgao")="87" then
                                       do until rs1.eof=true
                                       %>
                                       <option value="<%=rs1.fields(0).value%>"><%=rs1.fields(1).value%></option>   
                                       <%                                    
                                       rs1.movenext
                                       loop
                                       else
                                       do until rs1.eof=true
                                       set temp = db.execute("SELECT * FROM CURSO WHERE CURS_CD_CURSO='" & rs1.fields(0).value & "'")
                                       NOME_CURSO = temp("CURS_TX_NOME_CURSO")
                                       %>
                                       <option value="<%=rs1.fields(0).value%>"><%=NOME_CURSO%></option>   
                                       <%                                    
                                       rs1.movenext
                                       loop
                                       end if%>
                                       </select></td>
                         </tr>
                         <tr>
                                    <td width="100%" height="45" align="center" colspan="3"><p><a href="javascript:history.go(-1)"><img border="0" src="voltar.gif"></a>&nbsp;&nbsp; <a href="#" onClick="document.frm1.submit()"><img border="0" src="enviar.gif"></a></td>
                         </tr>
                         </table>
                      </td>
           </tr>
</table>
</form>
</body>

</html>

<script>
document.title = 'Indicação de Multiplicadores'
</script>