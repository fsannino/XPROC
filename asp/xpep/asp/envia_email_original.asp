<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")
	
strNomeOnda				= Request("pNomeOnda")	
str_Plano 				= Request("pPlano")
str_Problemas 			= Request("txtProblemas")
str_AcoesCorrConting 	= Request("txtAcoesCorrConting")
str_DTAprovacao_PAC 	= Request("txtDTAprovacao_PAC")
str_UsuarioResponsavel	= Request("txtUsuarioResponsavel")
str_RespTecSinGeral		= Request("txtRespTecSinGeral")
str_RespFunSinGeral		= Request("txtRespFunSinGeral")
str_Atividade			= Request("pNomeAtividade")
str_PlanoOrigem			= Request("pPlanoOrigem")

'Response.write "str_Plano " & str_Plano & "<br>"
'Response.write "str_Problemas " & str_Problemas 	& "<br>"
'Response.write "str_AcoesCorrConting " & str_AcoesCorrConting 	& "<br>"
'Response.write "str_DTAprovacao_PAC " & str_DTAprovacao_PAC & "<br>"
'Response.write "str_UsuarioResponsavel " & str_UsuarioResponsavel	& "<br>"
'Response.write "str_RespTecSinGeral " & str_RespTecSinGeral		& "<br>"
'Response.write "str_RespFunSinGeral " & str_RespFunSinGeral		& "<br>"
'Response.end

'*** USUARIO SINERGIA  - RESPONSÁVEL PELO PROCEDIMENTO***
'sql_RespProced= ""	
'sql_RespProced = sql_RespProced & " SELECT USMA_CD_USUARIO"		
'sql_RespProced = sql_RespProced & " FROM USUARIO_MAPEAMENTO "
'sql_RespProced = sql_RespProced & " WHERE USMA_TX_MATRICULA <> 0"
'sql_RespProced = sql_RespProced & " AND USMA_CD_USUARIO = '" & str_UsuarioResponsavel & "'"
'set rds_RespProced = db_Cogest.Execute(sql_RespProced)
'if not rds_RespProced.eof then
'	str_emailRespProced = """" & rds_RespProced("USUA_TX_EMAIL_EXTERNO") & """"
'	str_NomeRespProced  = rds_RespTecSin("USMA_TX_NOME_USUARIO")
'else
'	str_emailRespProced = ""	
'end if
'rds_RespProced.close
'set rds_RespProced = nothing
'Response.write str_emailRespProced

'*** USUARIO RESPONSÁVEL - SINERGIA TÉCNICO ***
sql_RespTecSin = ""	
sql_RespTecSin = sql_RespTecSin & " SELECT USUA_TX_NOME_USUARIO, USUA_TX_EMAIL_EXTERNO"		
sql_RespTecSin = sql_RespTecSin & " FROM USUARIO "
sql_RespTecSin = sql_RespTecSin & " WHERE USUA_CD_USUARIO = '" & str_RespTecSinGeral & "'"

set rds_RespTecSin = db_Cogest.Execute(sql_RespTecSin)

if not rds_RespTecSin.eof then
	str_emailRespTecSin = """" & rds_RespTecSin("USUA_TX_EMAIL_EXTERNO") & """"
	str_NomeRespTecSin  = rds_RespTecSin("USUA_TX_NOME_USUARIO")
else
	str_emailRespTecSin = ""	
end if
rds_RespTecSin.close
set rds_RespTecSin = nothing
'Response.write "Responsável Técnico - Sinergia - " & str_emailRespTecSin & "<br>" & str_RespFunSinGeral

'*** USUARIO RESPONSÁVEL - SINERGIA FUNCIONAL ***
sql_RespFunSin = ""	
sql_RespFunSin = sql_RespFunSin & " SELECT USUA_TX_NOME_USUARIO, USUA_TX_EMAIL_EXTERNO"		
sql_RespFunSin = sql_RespFunSin & " FROM USUARIO "
sql_RespFunSin = sql_RespFunSin & " WHERE USUA_CD_USUARIO = '" & str_RespFunSinGeral & "'"

set rds_RespFunSin = db_Cogest.Execute(sql_RespFunSin)

if not rds_RespFunSin.eof then
	str_emailRespFunSin = """" & rds_RespFunSin("USUA_TX_EMAIL_EXTERNO") & """"
	str_NomeRespFunSin  = rds_RespFunSin("USUA_TX_NOME_USUARIO")
else
	str_emailRespFunSin = ""		
end if
rds_RespFunSin.close
set rds_RespFunSin = nothing
'Response.write "Responsável Funcional - Sinergia - " & str_emailRespFunSin  & "<br>"

data_Atual = day(date) &"/"& month(date) &"/"& year(date)
str_txtEmail = ""
str_txtEmail = "Seguem as acões corretivas para o Plano - " & str_PlanoOrigem
str_txtEmail = str_txtEmail & ", pertencente a Onda - " & strNomeOnda 
str_txtEmail = str_txtEmail & " Acoes Corretivas - " & str_AcoesCorrConting
str_txtEmail = str_txtEmail & " Data: " & data_Atual 
str_txtEmail = str_txtEmail & " Enviado por: " & Session("CdUsuario")

intCountUsuarios = 0

set correio = Server.CreateObject("Persits.MailSender")
correio.host = "harpia.petrobras.com.br" 
correio.from = "acoescorretivas@cutover.com"

if str_emailRespProced <> "" then
	correio.AddAddress "rogerio.bl_informatica@petrobras.com.br" 'str_emailRespProced
	intCountUsuarios = intCountUsuarios + 1	
end if
if str_emailRespTecSin <> "" then 			
	correio.AddAddress "rogerio.bl_informatica@petrobras.com.br" 'str_emailRespTecSin
	intCountUsuarios = intCountUsuarios + 1
end if
if str_emailRespFunSin <> "" then
	correio.AddAddress "rogerio.bl_informatica@petrobras.com.br" 'str_emailRespFunSin
	intCountUsuarios = intCountUsuarios + 1	
end if

correio.Subject = "Açoes Corretivas"
correio.Body = str_txtEmail	
correio.send

if str_emailRespProced = "" or str_emailRespTecSin = "" or str_emailRespFunSin = "" then
		
	if intCountUsuarios > 1 then
		strMsg = "Não foi possível enviar o e-mail para o Usuário, " 
		strMsg = strMsg & "pois seu e-mail não está cadastrado.<br><br>"
	else
		strMsg = "Não foi possível enviar os e-mails para os Usuários, " 
		strMsg = strMsg & "pois seus e-mails não estão cadastrados.<br><br>"
	end if
		
	if str_emailRespProced = "" then	
		strMsg = strMsg & "Responsável pelo Procedimento  - " & str_NomeRespProced & " Chave: " & str_UsuarioResponsavel & "<br>"		
	end if
	if str_emailRespTecSin = "" then	
		strMsg = strMsg & "Responsável Técnico Sinergia   - " & str_NomeRespTecSin & " Chave: " & str_RespTecSinGeral & "<br>"		
	end if
	if str_emailRespFunSin = "" then			
		strMsg = strMsg & "Responsável Funcional Sinergia - " & str_NomeRespFunSin & " Chave: " & str_RespFunSinGeral & "<br>"
	end if	
else
	strMsg = "Os e-mails foram enviados com sucesso!"	
end if

Response.redirect "msg_geral.asp?pMsg=" & strMsg & "&pPlano=" & strPlano
%>