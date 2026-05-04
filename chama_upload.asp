<html>
	<head>
	    <%	
		strAcao = request("pAcao")		
		
		if strAcao = "IncluirArquivo" then	
			strTitulo = "Upload de Arquivo"	
			strOnLoad = "alert('Anexe um documento em formato doc com a sua Experiência Profissional.');"
		elseif strAcao = "AbrirArquivo" then
			strTitulo = "Abrir Arquivo"	
			strOnLoad = ""
		end if		
		%>	
		
		<title>:: <%=strTitulo%></title>
		
		<script language="JavaScript1.2">
			function Testando_Campo() 
			{
				var strAcao = document.frmUpload.pAcao.value;
				var resp = true;
				
				alert(strAcao);
				
				if (strAcao == 'IncluirArquivo')
				{										
					if (document.frmUpload.txtArquivo1.value.length == 0) 
					{
						alert("Você precisa selecionar algum arquivo para ser enviado.");
						resp=false;
					}
					
					var texto = document.frmUpload.txtArquivo1.value;
					var extensao = texto.substring(document.frmUpload.txtArquivo1.value.length-4,document.frmUpload.txtArquivo1.value.length);
					
					//if ((resp==true) && (document.frmUpload.NomeArqAtch.value.length != 0) && (extensao != '.doc')) 
					if ((resp==true) && (extensao != '.doc')) 
					{
						alert ('Somente podem ser enviados arquivos no formato DOC');
						resp=false;
					}	
					
					if (resp==true) 
					{				
						frmUpload.action='upload.asp?pAcao='+strAcao;
						document.frmUpload.submit();
					}
				}
				else
				{
					if (strAcao == 'AbrirArquivo')
					{
						if (document.frmUpload.txtArquivo1.value.length == 0) 
						{
							alert("Você precisa selecionar algum arquivo para ser enviado.");
							resp=false;
						}
						
						if (resp==true) 
						{				
							frmUpload.action='upload.asp?pAcao='+strAcao;
							document.frmUpload.submit();
						}
					}
				}
			}
		</script>	
	</head>
	
	<body onLoad="<%=strOnLoad%>">
		<form name="frmUpload" enctype="multipart/form-data" action="" method="post">
			<!--<input type="text" value="<%'=Session("Nome")%>" name="NomeArqAtch">	-->
			<input type="hidden" value="<%=strAcao%>" name="pAcao">				
			<center>
				<%
				if strAcao = "IncluirArquivo" then						
					%>
					<table border=0 align=center>
						<tr>		
							<td colspan="2" width=340><center><font size="2" face="verdana" id="menu"><b><%response.write ucase(strTitulo)%></b></font></td>				
						</tr>
						<tr>		
							<td colspan="2" width=340><font size="2" face="verdana"><b><%=strTitulo%></B></font><BR><small><font size=1 face=verdana>Procure o arquivo desejado:</small>			
								<input type="file" name="txtArquivo1" size="36"><br>
								<center><small><font size=1 face=verdana>[ Evite Fotografia, tamanho máximo permitido: 50kb ]</font></small><br>
							</td>		
						</tr>	
						<tr> 
						  <td align="left"> 
							<br>
							<a href="#" onClick='window.close();' id="menu"><small><font size=1 face=verdana>x Fechar</a>        
						  </td>
						  <td align="right">  
							<br>
							<a href="#" onClick='Testando_Campo();' id="menu"><small><font size=1 face=verdana>Enviar ></a>        
						  </td>
						</tr>		
					</table>
					<%					
				elseif strAcao = "AbrirArquivo" then
					%>						
					<table border=0 align=center>
						<tr>		
							<td colspan="2" width=340><center><font size="2" face="verdana" id="menu"><b><%=ucase(strTitulo)%></b></font></td>				
						</tr>
						<tr>		
							<td colspan="2" width=340><font size="2" face="verdana"><b><%=strTitulo%></B></font><BR><small><font size=1 face=verdana>Procure o arquivo desejado:</small>			
								<input type="file" name="txtArquivo1" size="36">								
							</td>		
						</tr>	
						<tr> 
						  <td align="left"> 
							<br>
							<a href="#" onClick='window.close();' id="menu"><small><font size=1 face=verdana>x Fechar</a>        
						  </td>
						  <td align="right">  
							<br>
							<a href="#" onClick='Testando_Campo();' id="menu"><small><font size=1 face=verdana>Enviar ></a>        
						  </td>
						</tr>		
					</table>				
					<%
				end if	
				%>	
		</center>
	</form>
</body>	
		
					
</html>