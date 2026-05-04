<!--#include file="conecta.asp" -->
<%
set objUSR = server.createobject("Seseg.Usuario")

chave = request("selFunc")
mega = request("selMega")

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
db2.CursorLocation=3

if request("tipo")=1 then
	tipo_cons="MULTIPLICADOR X CURSO"
	legenda = "SELECIONE O EMPREGADO DESEJADO"
	pagina = "consulta_mc.asp"
	set temp = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM CLI_ORGAO WHERE USMA_CD_USUARIO='" & chave & "'")
	if temp.eof=false then
		set rs = db.execute("SELECT DISTINCT USMA_CD_USUARIO, USMA_TX_NOME_USUARIO FROM USUARIO_MAPEAMENTO WHERE ORME_CD_ORG_MENOR LIKE '" & temp("ORME_CD_ORG_MENOR") & "%' AND USMA_CD_USUARIO<>'" & chave & "' ORDER BY USMA_TX_NOME_USUARIO")
	else
		if chave="SM23" then
			orgao_int = 55
		else
			if chave="DCX0" then
				orgao_int= 88
			else
				if chave="B511" then
					orgao_int = 14
				else
					if chave="EADE" or chave="RV61" then
						orgao_int=43
					else
						if chave="WS04" then
							orgao_int=87
						else
							orgao_int="XX"
						end if
					end if
				end if
			end if
		end if
		if orgao_int<>"XX" then		
			set rs = db.execute("SELECT DISTINCT USMA_CD_USUARIO, USMA_TX_NOME_USUARIO FROM USUARIO_MAPEAMENTO WHERE ORME_CD_ORG_MENOR LIKE '" & orgao_int & "%' AND USMA_CD_USUARIO<>'" & chave & "' ORDER BY USMA_TX_NOME_USUARIO")
		else
			set rs = db.execute("SELECT DISTINCT USMA_CD_USUARIO, USMA_TX_NOME_USUARIO FROM USUARIO_MAPEAMENTO WHERE PERF_CD_PERFIL<>2 AND USMA_CD_USUARIO<>'" & chave & "' ORDER BY USMA_TX_NOME_USUARIO")		
		end if
	
	end if
else
	tipo_cons="CURSO X MULTIPLICADOR"
	legenda = "SELECIONE O CURSO DESEJADO"
	pagina = "consulta_cm.asp"
	
	set is_cli = db.execute("SELECT * FROM CLI_ORGAO WHERE USMA_CD_USUARIO='" & chave & "'")
	
	if is_cli.eof=false then
		if len(is_cli("ORME_CD_ORG_MENOR"))=2 then
			org_cli = is_cli("ORME_CD_ORG_MENOR")
			set temp = db.execute("SELECT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & org_cli)
			tabela=temp("ORGAO")			
		else
			org_cli = left(is_cli("ORME_CD_ORG_MENOR"),7)
			set temp = db.execute("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & org_cli &"00000000' AND ORME_CD_STATUS='A'")
			tabela=temp("ORGAO")
		end if
	else
		tabela="XX"
	end if
	
	on error resume next
	ssql="SELECT DISTINCT CURSO FROM [" & tabela & "] ORDER BY CURSO"
	
	set rs = db2.execute(ssql)
	
	if rs.eof=true or err.number<>0 then
		if chave="SM23" then
			orgao_int2 = "ENGENHARIA"
		else
			if chave="DCX0" then
				orgao_int2= "AB"
			else
				if chave="B511" then
					orgao_int2 = "CENPES"
				else
					if chave="EADE" or chave="RV61" then
						orgao_int2="COMPARTILHADO"
					else
						if chave="WS04" then
							orgao_int2="E&P"
						else
							orgao_int2="XX"
						end if
					end if
				end if
			end if
		end if
		if orgao_int2<>"XX" then		
		
		ssql="SELECT DISTINCT CURSO FROM [" & ORGAO_INT2 & "] ORDER BY CURSO"
	
		on error resume next
		set rs = db2.execute(ssql)
		
		if err.number<>0 then
			set rs = db.execute("SELECT DISTINCT CURS_CD_CURSO AS CURSO FROM CURSO ORDER BY CURS_CD_CURSO")
			err.clear
		end if
		else
				set rs = db.execute("SELECT DISTINCT CURS_CD_CURSO AS CURSO FROM CURSO ORDER BY CURS_CD_CURSO")
		end if
		err.clear
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
                                    <td width="81%" height="24" align="left"><select size="1" name="selItem" style="font-family: Verdana; font-size: 7 pt">
                                       <option value="XXXX">== TODOS ==</option>
                                       <%
                                       do until rs.eof=true
                                       if request("tipo")=1 then
                                       %>
                                       <option value="<%=rs.fields(0).value%>"><%=rs.fields(1).value%></option>                                       
                                       <%
                                       else
									   set rs2=db.execute("SELECT DISTINCT CURS_TX_NOME_CURSO FROM CURSO WHERE CURS_CD_CURSO='" & rs("curso") & "'")									   
                                       if len(rs2.fields(1).value)>60 then
                                       %>
                                       <option value="<%=rs.fields(0).value%>"><%=left(rs2.fields(0).value,60)%>...</option>                                       
                                       <%
                                       else
                                       %>
                                       <option value="<%=rs.fields(0).value%>"><%=rs2.fields(0).value%></option>                                       
                                       <%
                                       end if
                                       end if
                                       rs.movenext
                                       loop
                                       %>
                                       </select></td>
                         </tr>
                         <tr>
                                    <td width="100%" height="45" align="center" colspan="3">&nbsp;<p><a href="javascript:history.go(-1)"><img border="0" src="voltar.gif"></a>&nbsp;&nbsp; <a href="#" onClick="document.frm1.submit()"><img border="0" src="enviar.gif"></a></td>
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