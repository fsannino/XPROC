<%@ Language=VBScript %> 
<%
strAcao = request("pAcao")

if strAcao = "IncluirArquivo" then
 
	Dim Contador, Tamanho 
	Dim ConteudoBinario, ConteudoTexto 
	Dim Delimitador, Posicao1, Posicao2 
	Dim ArquivoNome, ArquivoConteudo, PastaDestino 
	Dim objFSO, objArquivo 
	
	PastaDestino = replace(Request.ServerVariables("PATH_TRANSLATED"), Replace(Request.ServerVariables("PATH_INFO"),"/", "\"),"\") 	
	PastaDestino = PastaDestino & "xproc\Publico\"
	
	'***** Determina o Tamanho do Conteúdo ***** 
	Tamanho = Request.TotalBytes 
	
	'***** Obtém o Conteúdo no Formato Binário ***** 
	ConteudoBinario = Request.BinaryRead(Tamanho) 
	
	'***** Transforma o Conteúdo Binário em String ***** 
	For Contador = 1 To Tamanho 
		ConteudoTexto = ConteudoTexto & Chr(AscB(MidB(ConteudoBinario, Contador, 1))) 
	Next 
	
	'***** Determina o Delimitador de Campos ***** 
	Delimitador = Left(ConteudoTexto,InStr(ConteudoTexto, vbCrLf)-1) 
	
	'***** Percore a String Procurando os Campos ***** 
	'***** Identifica os Arquivos e Grava no Disco ***** 
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 
	
	Posicao1 = InStr(ConteudoTexto, Delimitador) + Len(Delimitador) 
	
	Do While True 
		ArquivoNome = "" 
		Posicao1 = InStr(Posicao1, ConteudoTexto, "filename=") 
		If Posicao1 = 0 Then 
			Exit Do 
		Else 
			'***** Determina o Nome do Arquivo ***** 
			Posicao1 = Posicao1 + 10 
			Posicao2 = InStr(Posicao1, ConteudoTexto, """") 
			For Contador = Posicao2-1 To Posicao1 Step -1 
				If Mid(ConteudoTexto, Contador, 1) <> "\" Then 
					ArquivoNome = Mid(ConteudoTexto, Contador, 1) & ArquivoNome 
				Else 
					Exit For 
				End If 
			Next 		
			
			'***** Determina o Conteúdo do Arquivo ***** 
			Posicao1 = InStr(Posicao1, ConteudoTexto, vbCrLf & vbCrLf) + 4 
			Posicao2 = InStr(Posicao1, ConteudoTexto, Delimitador) - 2 
			ArquivoConteudo = Mid(ConteudoTexto, Posicao1, Posicao2-Posicao1+1) 
															
			'ArquivoNome = Session("CdUsuario") & "_" & MontaDataHora(date(),2) & "_" & MontaDataHora(time(),3) 
			'Response.write ArquivoNome & "<br>"				
											
			'***** Grava o Arquivo ***** 
			If ArquivoNome <> "" Then 
				Set objArquivo = objFSO.CreateTextFile(PastaDestino & ArquivoNome, True) 
				objArquivo.WriteLine ArquivoConteudo
							 
				Response.Write "Arquivo " & PastaDestino & ArquivoNome & " gravado com sucesso!!!<BR>" 
				
				objArquivo.Close
				Set objArquivo = Nothing 
			End If 
		End If 
	Loop 
	
	Set objFSO = Nothing 
	
	public function MontaDataHora(strData,intDataTime)
	
		'*** intDataTime - Indica se mostraá a data c/ hora ou apenas a data.
		'*** intDataTime = 1 (DATA E HORA)
		'*** intDataTime = 2 (DATA)
		'*** intDataTime = 3 (HORA)
	
		if day(strData) < 10 then
			strDia = "0" & day(strData)		
		else
			strDia = day(strData)		
		end if
		
		if month(strData) < 10 then
			strMes = "0" & month(strData)	
		else
			strMes = month(strData)	
		end if		
		
		if hour(strData) < 10 then
			strHora = "0" & hour(strData)		
		else
			strHora = hour(strData)		
		end if
		
		if minute(strData) < 10 then
			strMinuto = "0" & minute(strData)	
		else
			strMinuto = minute(strData)	
		end if	
	
		if cint(intDataTime) = 1 then	
			MontaDataHora = strDia & "/" & strMes & "/" & year(strData) & " - " &  strHora & ":" & strMinuto	
		elseif cint(intDataTime) = 2 then	
			MontaDataHora = strDia & "_" & strMes & "_" & year(strData) 
		elseif cint(intDataTime) = 3 then	
			MontaDataHora = strHora & "_" & strMinuto	
		end if
	end function
elseif  strAcao = "AbrirArquivo" then
									
	PastaDestino = replace(Request.ServerVariables("PATH_TRANSLATED"), Replace(Request.ServerVariables("PATH_INFO"),"/", "\"),"\") 	
	PastaDestino = PastaDestino & "xproc\Publico\"
		
	ArquivoNome = "testfile.doc"
		
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject") 					
	Set txtfile = objFSO.CreateTextFile(PastaDestino & ArquivoNome, true)
		
	'*** Write a line.
	Set fil1 = objFSO.GetFile(PastaDestino & ArquivoNome)
	Set ts = fil1.OpenAsTextStream(1,2)	
	'Set ts = fil1.OpenAsTextStream(ForWriting)
	'ts.Write "Hello Mundo"
	'ts.Close
	
	'*** Lê o conteúdo do arquivo.
	'Set ts = fil1.OpenAsTextStream(2,2)
	'Set ts = fil1.OpenAsTextStream(ForReading)
	's = ts.ReadAll
	
	'Response.write s
	
	'ts.Close				
end if
%> 
