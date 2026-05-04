<%
Session("Corte") 		= request("selCorte")

strUnidade 				= request("selUnidade")
strDiretoria 			= request("selDiretoria")
strCurso 				= request("selCurso")

strDescentralizado 		= request("rdDescentralizado")
strEaD 					= request("rdEad")
strInLoco 				= request("rdInLoco")

strNumRel				= request("pNumRel")
strTituloRel			= request("pTituloRel")

'Response.write "strNumRel - " & strNumRel & "<br>"
'Response.write "strTituloRel - " & strTituloRel & "<br>"

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("Petrobras 2004_v2.mdb")
'db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

if request("excel") = 1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if

strSQLTurma = ""
strSQLTurma = strSQLTurma & "SELECT Diretoria, Unidade, [Centro de Treinamento], "
strSQLTurma = strSQLTurma & "[Quantidade de Usuários], CodCurso, Turma, Sala, [Data de Início], " 
strSQLTurma = strSQLTurma & "[Data de Término], Multiplicador, Descentralizado, EaD, [In loco] "
strSQLTurma = strSQLTurma & "FROM [Demanda x Oferta Geral] "
strSQLTurma = strSQLTurma & "WHERE Diretoria  ='" & replace(strDiretoria,"_","&") & "' "

if strUnidade <> "0" then
	strSQLTurma = strSQLTurma & " AND Unidade ='" & strUnidade & "' "
end if 

if strCurso <> "0" then
	strSQLTurma = strSQLTurma & " AND CodCurso ='" & strCurso & "' "
end if 

if strEaD = "0" then
	strSQLTurma = strSQLTurma & "And EaD = True "
elseif strEaD = "1" then
	strSQLTurma = strSQLTurma & "And EaD = False "
end if

if strDescentralizado = "0" then
	strSQLTurma = strSQLTurma & "And Descentralizado = False "
elseif strDescentralizado = "1" then
	strSQLTurma = strSQLTurma & "And Descentralizado = True "
end if

if strInLoco = "0" then
	strSQLTurma = strSQLTurma & "AND [In loco] = True "
elseif strInLoco = "1" then
	strSQLTurma = strSQLTurma & "AND [In loco] = False "
end if
		
'Response.write strSQLTurma
'Response.end		
		
set rstTurmas = db_banco.execute(strSQLTurma)				
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		<style type="text/css">
			<!--
			.style2 
			{
				font-family: Verdana, Arial, Helvetica, sans-serif;
				font-weight: bold;
				color: #000066;
			}		
			-->
		</style>
	</head>
	
	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">		
		<%		
		if request("excel") <> 1 then
		%>		
		<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
		  <tr>
			<td width="20%" height="20">&nbsp;</td>
			<td width="44%" height="60">&nbsp;</td>
			<td width="36%" valign="top"> 
			  <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
				<tr> 
				  <td bgcolor="#330099" width="39" valign="middle" align="center"> 
					<div align="center">
					  <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" width="27" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif"></a></div>
				  </td>
				</tr>
				<tr> 
				  <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
					<div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
					<div align="center"><a href="../../indexA_grade.asp?selCorte=<%=Session("Corte")%>"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
				  </td>
				</tr>
			  </table>
			</td>
		  </tr>
		  <tr bgcolor="#F1F1F1">
			<td colspan="3" height="20">
			  <table width="625" border="0" align="center">
				<tr>
					<td width="26"></td>
				  <td width="50"></td>
				  <td width="26">&nbsp;</td>
				  <td width="195"></td>
					<td width="27"></td>  
					<td width="50"></td>
				  <td width="28"></td>
				  <td width="26">&nbsp;</td>
				  <td width="159"></td>
				</tr>
			  </table>
			</td>
		  </tr>
		</table>	
		
			<table cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td width="720"></td>	
					<td width="50">
						<div align="center">	
							 <a href="gera_consulta_turma.asp?excel=1&amp;selDiretoria=<%=strDiretoria%>&amp;selUnidade=<%=strUnidade%>&amp;rdDescentralizado=<%=strDescentralizado%>&amp;rdEad=<%=strEaD%>&amp;rdInLoco=<%=strInLoco%>&amp;selCurso=<%=strCurso%>&amp;pTituloRel=<%=strTituloRel%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif" title="Exportar para o Excel"></a>
						</div>
					</td>						
				</tr>
			</table>
		<%end if%>		
		
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
			<tr>
			  <td height="10">
			  </td>
			</tr>
			<tr>
			  <td>
				<div align="center"><font face="Verdana" color="#330099" size="3"><b>Relatório de <%=strTituloRel%> - Grade de Treinamento</b></font></div>
			  </td>
			</tr>
			<!--	<tr>
			  <td>&nbsp;</td>
		  </tr>
		  		
		    
		  <tr>
			  <td>
				  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
					<tr>
					  <td width="1%">&nbsp;</td>
					  <td width="9%"><font face="Verdana" color="#330099" size="2"><b>Diretoria:</b></font></td>
					  <td width="90%"><font face="Verdana" color="#330099" size="2"><%'=Ucase(replace(strDiretoria,"_","&"))%></font></td>
					</tr>
				  </table>
			  </td>
		  </tr>
		  -->
		  
		  <%'if strUnidade <> "0" then%>
		  	 <!-- <tr>
				  <td>
					  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
						<tr>
						  <td width="1%">&nbsp;</td>
						  <td width="9%"><font face="Verdana" color="#330099" size="2"><b>Unidade:</b></font></td>
						  <td width="90%"><font face="Verdana" color="#330099" size="2"><%'=Ucase(strUnidade)%></font></td>
						</tr>
					  </table>
				  </td>
			  </tr>-->
		  <%'end if%>
			<tr>
			  <td>&nbsp;</td>
			</tr>
    </table>		
				
		<table width="1001" border="0" cellpadding="2" cellspacing="2">
		  <tr bgcolor="#D4D0C8">
		    <td width="58" align="center">
					<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Diretoria</b></font>
			</td>
			<%'if strUnidade = "0" then%>
				<td width="81" align="center">
					<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Unidade</b></font>
				</td>
			<%'end if%>									
		    <td width="91" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Centro de Treinamento</b></font>
			</td>		  
		     <td width="50" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Cód Curso</b></font>
			</td>
			<td width="46" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Qtde de Usuários</b></font>
			</td>		   
		    <td width="206" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Turma</b></font>
			</td>
		    <td width="92" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Sala</b></font>
			</td>
		    <td width="32" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Data de Início</b></font>
			</td>		   
			<td width="44" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Data de T&eacute;rmino</b></font>
			</td>       
			<td width="102" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Multiplicador</b></font>
			</td>			
			<td width="139" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Tipo de Curso</b></font>
			</td>	
			<!--		
			<td width="27" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Ead</b></font>
			</td>	
			<td width="48" align="center">
				<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>InLoco</b></font>
			</td>		
			-->
		  </tr>		
			<%
			'*** INICIALIZAÇÕES ***	
			strCor = "#FFFFFF"
			intTotal = 0			
			strNomeAtual = ""
			strNomeCurso = ""
			intQtdeUsuario = 0
			intTotalQtdeUsuario = 0
			
			if not rstTurmas.eof then									
				do until rstTurmas.eof = true
														
					if strCor = "#FFFFFF" then
						strCor = "#EAEAEA"
					else
						strCor = "#FFFFFF"
					end if
					
					if strNomeAtual = "" then
						strNomeCurso = trim(rstTurmas("CodCurso"))
						strNomeAtual = strNomeCurso
						
						intQtdeUsuario = rstTurmas("Quantidade de Usuários")
						intTotalQtdeUsuario = intTotalQtdeUsuario + rstTurmas("Quantidade de Usuários")						
					else
						if strNomeAtual = trim(rstTurmas("CodCurso")) then
							'strNomeCurso = ""
							intQtdeUsuario = ""
						else
							strNomeCurso = trim(rstTurmas("CodCurso"))
							strNomeAtual = strNomeCurso
							
							intQtdeUsuario = rstTurmas("Quantidade de Usuários")
							intTotalQtdeUsuario = intTotalQtdeUsuario + rstTurmas("Quantidade de Usuários")						
						end if						
					end if															
					%>					
					<tr bgcolor="<%=strCor%>">
						<td width="58" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rstTurmas("Diretoria")%></font>
						</td>
						<%'if strUnidade = "0" then%>
							<td width="81" align="center">
								<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rstTurmas("Unidade")%></font>
							</td>
						<%'end if%>								
						<td width="91" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rstTurmas("Centro de Treinamento")%></font>
						</td>
						<td width="50" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strNomeCurso%></font>
						</td>						
						<td width="46" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=intQtdeUsuario%></font>
						</td>						
						<td width="206" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rstTurmas("Turma")%></font>
						</td>
						<td width="92" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rstTurmas("Sala")%></font>
						</td>
						
						<% 	
						if Trim(rstTurmas("Data de Início")) <> "" then
							strDia = ""		
							strMes = ""
							strAno = ""
							vetDtAprov = split(Trim(rstTurmas("Data de Início")),"/")						
							strDia = trim(vetDtAprov(0))
							if cint(strDia) < 10 then
								strDia = "0" & strDia
							end if			
							strMes = trim(vetDtAprov(1))			
							if cint(strMes) < 10 then
								strMes = "0" & strMes
							end if
							strAno = trim(vetDtAprov(2))
							dat_DtInicio = strDia & "/" & strMes & "/" & strAno 
						else
							dat_DtInicio = ""
						end if
						%>
						<td width="32" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=dat_DtInicio%></font>
						</td>	
						
						<% 	
						if Trim(rstTurmas("Data de Término")) <> "" then
							strDia = ""		
							strMes = ""
							strAno = ""
							vetDtAprov = split(Trim(rstTurmas("Data de Término")),"/")						
							strDia = trim(vetDtAprov(0))
							if cint(strDia) < 10 then
								strDia = "0" & strDia
							end if			
							strMes = trim(vetDtAprov(1))			
							if cint(strMes) < 10 then
								strMes = "0" & strMes
							end if
							strAno = trim(vetDtAprov(2))
							dat_DtFim = strDia & "/" & strMes & "/" & strAno 
						else
							dat_DtFim = ""
						end if
						%>					
						<td width="44" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=dat_DtFim%></font>
						</td>					    
						<td width="102" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rstTurmas("Multiplicador")%></font>
						</td>								
						<%
						'if rstTurmas("EaD") = True then
						'	strTxtEAD = "Sim"
						'else
						'	strTxtEAD = "Não"
						'end if	
													
						'if rstTurmas("Tipo do Curso") = True then
						'	strTxtCentralizada = "Descentralizado"
						'else
						'	strTxtCentralizada = "Centralizado"
						'end if												
												
						'if rstTurmas("In loco") = True then
						'	strTxtInLoco = "Sim"
						'else
						'	strTxtInLoco = "Não"
						'end if
						
						strTipoCurso = "" 
						
						if rstTurmas("EaD") = True then
							strTipoCurso = "EAD"
						else
							strTipoCurso = "Presencial"
						end if	
													
						if rstTurmas("Descentralizado") = True then
							strTipoCurso = strTipoCurso & " Descentralizado"
						else
							strTipoCurso = strTipoCurso & " Centralizado"
						end if												
												
						if rstTurmas("In loco") = True then
							strTipoCurso = strTipoCurso & " In loco"						
						end if						
						%>									
						<td width="139" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strTipoCurso%></font>
						</td>						
						
						<!--
						<td width="27" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%'=strTxtEAD%></font>
						</td>												
							
						<td width="48" align="center">
							<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%'=strTxtInLoco%></font>
						</td>			
						-->			
					</tr>					
					<%
					intTotal = intTotal + 1					
					rstTurmas.movenext					
				loop
				%>
				<tr><td height="20"></td></tr>
				<tr>
					<td colspan="10">
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Total de Turmas:</b></font>&nbsp;&nbsp;  
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><%=intTotal%></b></font>				  
					</td>
				</tr>
				<tr>				
					<td colspan="10">
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Total de Usuários:</b></font>&nbsp;&nbsp;  
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><%=intTotalQtdeUsuario%></b></font>				  
					</td>					
				</tr>				
				<tr><td height="10"></td></tr>
				<%
			else
			%>
				<tr>
					<td colspan="10">
						<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1">Não existe registros para esta consulta.</font>
					</td>
				</tr>					
			<%
			end if
			
			rstTurmas.close
			set rstTurmas = nothing			
			%>
	</table>		
	</body>	
	<%
	db_banco.close
	set db_banco = nothing
	%>		
</html>
