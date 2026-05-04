<%
Dim iRet  ' retorno da função Envia_Emasil  = (-1) - Erro (1) - Enviado OK

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
'Response.write "str_Atividade " & str_Atividade		& "<br>"
'Response.write "str_PlanoOrigem " & str_PlanoOrigem		& "<br>"
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

'*** USUARIO REMETENTE - SINERGIA TÉCNICO ***
sql_Remetente = ""	
sql_Remetente = sql_Remetente & "SELECT USUA_TX_NOME_USUARIO, USUA_TX_EMAIL "		
sql_Remetente = sql_Remetente & "FROM XPEP_EQUIPE_SINERGIA "
sql_Remetente = sql_Remetente & "WHERE USUA_TX_CD_USUARIO = '" & Session("CdUsuario") & "'"

set rds_Remetente = db_Cogest.Execute(sql_Remetente)

if not rds_Remetente.eof then
	str_NomeRemetente  = rds_Remetente("USUA_TX_NOME_USUARIO")
else
	str_NomeRemetente = ""	
end if
rds_Remetente.close
set rds_Remetente = nothing

'*** USUARIO RESPONSÁVEL - SINERGIA TÉCNICO ***
sql_RespTecSin = ""	
sql_RespTecSin = sql_RespTecSin & "SELECT USUA_TX_NOME_USUARIO, USUA_TX_EMAIL "	
sql_RespTecSin = sql_RespTecSin & "FROM XPEP_EQUIPE_SINERGIA "
sql_RespTecSin = sql_RespTecSin & "WHERE USUA_TX_CD_USUARIO = '" & str_RespTecSinGeral & "'"

set rds_RespTecSin = db_Cogest.Execute(sql_RespTecSin)

if not rds_RespTecSin.eof then
	'str_emailRespTecSin = """" & rds_RespTecSin("USUA_TX_EMAIL_EXTERNO") & """"
	str_NomeRespTecSin  = rds_RespTecSin("USUA_TX_NOME_USUARIO")
else
	str_emailRespTecSin = ""	
end if
rds_RespTecSin.close
set rds_RespTecSin = nothing
'Response.write "Responsável Técnico - Sinergia - " & str_emailRespTecSin & "<br>" & str_RespFunSinGeral

'*** USUARIO RESPONSÁVEL - SINERGIA FUNCIONAL ***
sql_RespFunSin = ""	
sql_RespFunSin = sql_RespFunSin & "SELECT USUA_TX_NOME_USUARIO, USUA_TX_EMAIL "			
sql_RespFunSin = sql_RespFunSin & " FROM XPEP_EQUIPE_SINERGIA "
sql_RespFunSin = sql_RespFunSin & " WHERE USUA_TX_CD_USUARIO = '" & str_RespFunSinGeral & "'"

set rds_RespFunSin = db_Cogest.Execute(sql_RespFunSin)

if not rds_RespFunSin.eof then
	'str_emailRespFunSin = """" & rds_RespFunSin("USUA_TX_EMAIL_EXTERNO") & """"
	str_NomeRespFunSin  = rds_RespFunSin("USUA_TX_NOME_USUARIO")
else
	str_emailRespFunSin = ""		
end if
rds_RespFunSin.close
set rds_RespFunSin = nothing
'Response.write "Responsável Funcional - Sinergia - " & str_emailRespFunSin  & "<br>"

data_Atual = day(date) &"/"& month(date) &"/"& year(date)
str_txtEmail = ""
str_txtEmail = "Seguem as acões corretivas para a Atividade " & str_PlanoOrigem 
str_txtEmail = str_txtEmail & ", pertencente a Onda - " & strNomeOnda 
str_txtEmail = str_txtEmail & ", com o seguinte problema - " & str_Problemas
str_txtEmail = str_txtEmail & " e Acoes Corretivas - " & str_AcoesCorrConting
str_txtEmail = str_txtEmail & "       Data: " & data_Atual 
str_txtEmail = str_txtEmail & "       Enviado por: " & Session("CdUsuario")
'Response.Write str_txtEmail
'Response.end

intCountUsuarios = 0

'set correio = Server.CreateObject("Persits.MailSender")
'correio.host = "harpia.petrobras.com.br" 
'correio.from = "acoescorretivas@cutover.com"

'PROCEDIMENTO
'str_UsuarioResponsavel = ""
if str_UsuarioResponsavel <> "" then
	sNomeRemetente = str_NomeRemetente
	sRemetente =  Session("CdUsuario") & "S600146.petrobras.com.br"
	sDestinatario = str_UsuarioResponsavel & "@petrobras.com.br"
	sSubjectEmail = "Açoes Corretivas"
	sCorpoEmail = str_txtEmail
	sMsgErro =  "Erro ao enviar e-mail:"
	iRet=EnviarEmail(sNomeRemetente,sRemetente, "", sDestinatario, sSubjectEmail, sCorpoEmail, "", sMsgErro, "", "")
	if iRet = -1 then
   	   str_emailRespProced = ""
	else
	   str_emailRespProced = "Ok"
       intCountUsuarios = intCountUsuarios + 1		   
	end if   	
end if

'str_RespTecSinGeral = ""
if str_RespTecSinGeral <> "" then 			
	sNomeRemetente = str_NomeRemetente
	sRemetente =  Session("CdUsuario") & "S600146.petrobras.com.br"
	sDestinatario = str_RespTecSinGeral & "@petrobras.com.br"
	sSubjectEmail = "Açoes Corretivas"
	sCorpoEmail = str_txtEmail
	sMsgErro =  "Erro ao enviar e-mail:"
	iRet=EnviarEmail(sNomeRemetente,sRemetente, "", sDestinatario, sSubjectEmail, sCorpoEmail, "", sMsgErro, "", "")
	if iRet = -1 then
   	   str_emailRespTecSin = ""
	else
	   str_emailRespTecSin = "Ok"
	   intCountUsuarios = intCountUsuarios + 1
	end if   	
end if

'str_RespFunSinGeral = "xt54"
if str_RespFunSinGeral <> "" then
	sNomeRemetente = str_NomeRemetente
	sRemetente =  Session("CdUsuario") & "S600146.petrobras.com.br"
	sDestinatario = str_RespFunSinGeral & "@petrobras.com.br"
	sSubjectEmail = "Açoes Corretivas"
	sCorpoEmail = str_txtEmail
	sMsgErro =  "Erro ao enviar e-mail:"
	iRet=EnviarEmail(sNomeRemetente,sRemetente, "", sDestinatario, sSubjectEmail, sCorpoEmail, "", sMsgErro, "", "")
	if iRet = -1 then
   	   str_emailRespFunSin = ""
	else
	   str_emailRespFunSin = "Ok"
	   intCountUsuarios = intCountUsuarios + 1	
	end if   	
end if

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

'-----------------------------------------------------------------------------
' Função EnviarEmail( 	BYVAL asNmRemetente, BYVAL asEmailRemetente			_
'								BYVAL asNmDestinatario, BYVAL asEmailDestinatario, _
'								BYVAL asSubject, BYVAL asCorpoEmail, 					_
'								BYVAL avsEmailBcc, BYREF asMsgErro 						)
'
'		Função usada para enviar email.
'
'		Recebe:
'			String	asNmRemetente			Nome do remetente
'			String	asEmailRemetente		Email do remetente
'			String	asNmDestinatario		Nome do destinatário
'			String	asEmailDestinatario		Email do destinatário
'			String	asSubject				Subject da mensagem
'			String	asCorpoEmail			Corpo da mensagem
'			VetStr	avsEmailBcc				Email dos usuário que receberão via bcc
'			String	asMsgErro				Mensagem de erro
'			String	asAnexo					Caminho do Arquivo a ser anexado
'
'		Retorna: -1 Erro
'					 1 Ok
'
'	Autor: Alexandre Motta Drummond
'
'	Data: 30/07/2002
'
'-----------------------------------------------------------------------------
FUNCTION EnviarEmail( 	BYVAL asNmRemetente, BYVAL asEmailRemetente,				_
								BYVAL asNmDestinatario, BYVAL asEmailDestinatario, 	_
								BYVAL asSubject,  BYVAL asCorpoEmail, 						_
								BYVAL avsEmailBcc(), BYREF asMsgErro, ByVal asCC, BYVAL asAnexo)

	DIM oRs			' RecordSet
	DIM iRet			' Retorno de chamadas de funções
	DIM Mail, i, sNome

	On Error Resume Next

	'##MEXER AQUI para alterar o mail host##
	Dim  sTPCOMPONENTE_EMAIL
	DIm  sEMAIL_HOST
	
	sTPCOMPONENTE_EMAIL = 1
	sEMAIL_HOST = "164.85.62.165"

		' Criando o objeto de email
		Set Mail = Server.CreateObject("Persits.MailSender")

		' Preparando o Email
		Mail.Host		= sEMAIL_HOST	   	' Especificando um  SMTP server válido
		Mail.From		= asEmailRemetente	' Especificando o email de quem está enviando
		Mail.FromName		= asNmRemetente		' Especificando o nome de quem está enviando
		
		if asCC <> "" then 
			Mail.AddCC asCC, ""
		end if
		
		IF  asEmailDestinatario <> "" AND Not isNull(asEmailDestinatario)  THEN
			Dim  MyArray
			MyArray = Split(asEmailDestinatario, ",", -1, 1)

			FOR i = 0 TO UBOUND(MyArray)

				sNome = "Chave:"&MyArray(i)

				Mail.AddAddress Trim(MyArray(i)), sNome  'Especificando o endereço e nome do destinatário
			NEXT
		END IF
		
		Mail.Subject		= asSubject	' Especificando o subject do email
		Mail.IsHtml		= False		' Especificando que o email não será enviado no formato Html
		Mail.Body		= asCorpoEmail	' Especificando o corpo do email
		
		if asAnexo <>  "" then
			Mail.AddAttachment asAnexo
		end if
		
		' Acrescentando os Bcc´s
		IF  avsEmailBcc <> "" AND  Not isNull(avsEmailBcc)  THEN
			FOR i = 1 TO UBOUND( avsEmailBcc )
				Mail.AddBcc	avsEmailBcc( i )
			NEXT
		END IF	

		' Enviando o email e tratando o retorno
		Mail.Send
		IF Err.Number <> 0 THEN
			asMsgErro = "Erro ao enviar e-mail: " & Err.Description
			EnviarEmail = -1
			EXIT FUNCTION
		END IF
		' Destruindo o Objeto de email
		SET Mail = NOTHING

		SET Mail = NOTHING

	' Foi tudo bem então retorno 1.
	EnviarEmail = 1

	On Error Goto 0

END FUNCTION

%>