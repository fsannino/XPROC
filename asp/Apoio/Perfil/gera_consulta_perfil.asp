<!--#include file="conn_consulta.asp" -->
<html>
<%
server.ScriptTimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.cursorlocation=3

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

ssql=""
ssql="SELECT DISTINCT dbo.USUARIO_PERFIL.USPE_CD_USUARIO, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, "
ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR, dbo.USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
ssql=ssql+"FROM dbo.USUARIO_PERFIL "
ssql=ssql+"INNER JOIN dbo.USUARIO_MAPEAMENTO ON "
ssql=ssql+"dbo.USUARIO_PERFIL.USPE_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
ssql=ssql+"INNER JOIN dbo.ORGAO_MENOR ON "
ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR "
ssql=ssql+"INNER JOIN dbo.MACRO_PERFIL ON "
ssql=ssql+"dbo.USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL WHERE "

if len(orgao)>0 then
	ssql = ssql +  "dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & orgao & "%' " 
	tem_o = 1
end if
if request("selFuncao")<>"N" then
	ssql = ssql +  "AND dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = '" & request("selFuncao") & "' " 
else
if request("caso") = 1 then
	ssql = ssql +  "AND dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO LIKE 'HR.%' "
else
	ssql = ssql +  "AND dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO LIKE 'MM.%' "
end if	
end if

ssql=ssql+" order by dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "

set rs=db.execute(ssql)

if rs.eof=true then

	ssql=""
	ssql="SELECT DISTINCT dbo.USMA_MICRO_R3_VISAO_R3.USMA_CD_USUARIO, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR, dbo.USMA_MICRO_R3_VISAO_R3.MCPR_NR_SEQ_MACRO_PERFIL "
	ssql=ssql+"FROM dbo.USMA_MICRO_R3_VISAO_R3 "
	ssql=ssql+"INNER JOIN dbo.USUARIO_MAPEAMENTO ON "
	ssql=ssql+"dbo.USMA_MICRO_R3_VISAO_R3.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
	ssql=ssql+"INNER JOIN dbo.ORGAO_MENOR ON "
	ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR "
	ssql=ssql+"INNER JOIN dbo.MACRO_PERFIL ON "
	ssql=ssql+"dbo.USMA_MICRO_R3_VISAO_R3.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL WHERE "

	if len(orgao)>0 then
		ssql = ssql +  "dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & orgao & "%' " 
	tem_o = 1
	end if
	if request("selFuncao")<>"N" then
		ssql = ssql +  "AND dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = '" & request("selFuncao") & "' " 
	else
	if request("caso") = 1 then
		ssql = ssql +  "AND dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO LIKE 'HR.%' "
	else
		ssql = ssql +  "AND dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO LIKE 'MM.%' "
	end if	
	end if

	ssql=ssql+" order by dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USMA_MICRO_R3_VISAO_R3.MCPR_NR_SEQ_MACRO_PERFIL "

	set rs=db.execute(ssql)

end if

tem = rs.RecordCount
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Usuários cadastrados com Perfil</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>

<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#E5E5E5">
<form name="frm1">
<table width="81%" height="26" border="0">
           <tr>
                      <td width="5%">
                         <div align="right">
                                   <a href="javascript:history.go(-1)"><img src="seta_esquerda_01.jpg" width="21" height="18" border="0"></a></div>
                      </td>
                      <td width="5%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Voltar</font></strong></td>
                      <td width="4%">
                         <div align="right">
                                   <%if tem>0 then%>
                                   <a href="javascript:print()"><img src="impressão.jpg" width="27" height="21" border="0">
                                   </a>
                                   <%end if%>
                                   </div>
                      </td>
                      <td width="9%">
                      <%if tem>0 then%>
                      <strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Imprimir</font></strong>
                      <%end if%>
                      </td>
                      <td width="22%">&nbsp;</td>
           </tr>
</table>
<table width="93%" border="0">
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<table border="1" cellpadding="0" cellspacing="0" style="border-width:0; border-collapse: collapse" bordercolor="#111111" width="93%" id="AutoNumber1" height="47">
           <%
           if tem>0 then
           %>
           <tr>
                      <td width="4%" style="border-style: none; border-width: medium" height="17"></td>
                      <td width="61%" style="border-style: none; border-width: medium" bgcolor="#E5E5E5" height="17" colspan="3"><b><font face="Verdana" size="2" color="#000080">Usuários Cadastrados no R/3 com </font></b><font face="Verdana" color="#000080"><b><font size="2">Perfil</font></b></font></td>
                      <td width="48%" style="border-style: none; border-width: medium" bgcolor="#E5E5E5" height="17">&nbsp;</td>
           </tr>
           <tr>
                      <td width="4%" style="border-style: none; border-width: medium" height="17"></td>
                      <td width="8%" style="border-style: none; border-width: medium" bgcolor="#E5E5E5" height="17">&nbsp;</td>
                      <td width="30%" style="border-style: none; border-width: medium" bgcolor="#E5E5E5" height="17">&nbsp;</td>
                      <td width="23%" style="border-style: none; border-width: medium" bgcolor="#E5E5E5" height="17">&nbsp;</td>
                      <td width="48%" style="border-style: none; border-width: medium" bgcolor="#E5E5E5" height="17">&nbsp;</td>
           </tr>
           <tr>
                      <td width="4%" style="border-style: none; border-width: medium" height="17"></td>
                      <td width="8%" style="border-style: none; border-width: medium" bgcolor="#C0C0C0" height="17"><b><font face="Verdana" size="2" color="#000080">Chave</font></b></td>
                      <td width="30%" style="border-style: none; border-width: medium" bgcolor="#C0C0C0" height="17"><b><font face="Verdana" size="2" color="#000080">Nome</font></b></td>
                      <td width="23%" style="border-style: none; border-width: medium" bgcolor="#C0C0C0" height="17"><b><font face="Verdana" size="2" color="#000080">Lotação</font></b></td>
                      <td width="48%" style="border-style: none; border-width: medium" bgcolor="#C0C0C0" height="17"><b><font face="Verdana" size="2" color="#000080">Função</font></b></td>
           </tr>
		           <%
		           end if
					
					chave_ant = ""
					
					i = 0
										
					do until i = tem
					
					chave_atual = rs.fields(0).value
					nome = rs("USMA_TX_NOME_USUARIO")
					lotacao = rs("ORME_SG_ORG_MENOR")
					
					if chave_atual = chave_ant then
						chave_atual = " "
						nome = " "
						lotacao = " "					
					end if

					if cor="white" then
						cor="#E5E5E5"
					else
						cor="white"
				   end if			

                   ssql1=""
                   ssql1="SELECT * FROM MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL= " & RS("MCPR_NR_SEQ_MACRO_PERFIL")                     
                   set rs1 = db.execute(ssql1)
                   
                   if rs1.eof=false then
		           %>
		           <tr>
                      <td width="4%" style="border-style: none; border-width: medium" height="26">&nbsp;</td>
                      <td width="8%" style="border-style: none; border-width: medium" height="26" bgcolor="<%=cor%>"><font face="Verdana" size="1"><b><%=chave_atual%></b></font>&nbsp;</td>
                      <td width="30%" style="border-style: none; border-width: medium" height="26" bgcolor="<%=cor%>"><font face="Verdana" size="1"><b><%=nome%></b></font>&nbsp;</td>
                      <td width="23%" style="border-style: none; border-width: medium" height="26" bgcolor="<%=cor%>"><font face="Verdana" size="1"><%=lotacao%></font>&nbsp;</td>
                      <%
						ssql=""                      
						ssql="SELECT FUNE_TX_TITULO_FUNCAO_NEGOCIO FROM FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & RS1("FUNE_CD_FUNCAO_NEGOCIO") & "'"
						
						set temp = db.execute(ssql)
						
						titulo = temp("FUNE_TX_TITULO_FUNCAO_NEGOCIO")  
                          
                            IF RIGHT(TRIM(TITULO),6) = "ANTIGA" THEN
                            	TITULO = LEFT(TITULO,LEN(TITULO)-7)
                            	TITULO = TITULO & " ANTECIP"
                            ELSE
                            	TITULO=TITULO
                            END IF
                    
                      %>
                      <td width="48%" style="border-style: none; border-width: medium" height="26" bgcolor="<%=cor%>"><font face="Verdana" size="1"><b><%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%></b> - <%=TITULO%></font>&nbsp;</td>
		           </tr>
		           <%
		           end if
		           chave_ant = rs.fields(0).value
		           i=i+1
		           rs.movenext
		           loop
		           %>
</table>
<%if tem < 1 then%>
<p><font color="#800000"><strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Nenhum Registro Encontrado para a Seleção</strong></font></p>
<%end if%>
</form>
</body>

</html>