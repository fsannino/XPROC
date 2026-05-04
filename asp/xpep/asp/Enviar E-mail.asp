<%  

João,

Segue em anexo a função que faz o envio de e-mails. A chamada da mesma é feita da seguinte forma:

iRet=EnviarEmail(sNomeRemetente, sRemetente, "", sDestinatario , sSubjectEmail, sCorpoEmail, "", sMsgErro, "" , "") 

Onde:
O remetente deve estar obrigatóriamente no formato:
sRemetente = <CHAVE> & "@S600146.petrobras.com.br" 
E o destinatário deverá estar no formato:
sDestinatario = <CHAVE> & "@petrobras.com.br" 


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