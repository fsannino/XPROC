<%	
if trim(Session("Conn_String_Cogest_Gravacao")) = "" then
	Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest" 
end if

set db_cogest = Server.CreateObject("ADODB.Connection")
db_cogest.Open Session("Conn_String_Cogest_Gravacao")

strCase = trim(Request("selCase"))

on error resume next

sqlCase = ""
sqlCase = sqlCase & "SELECT CASE_TX_CD_CASE, CASE_NR_CD_CONDICAO, "
sqlCase = sqlCase & "TRAN_CD_TRANSACAO, CASE_TX_DESC_CASE, "
sqlCase = sqlCase & "CASE_TX_DESC_CONDICAO, CASE_DT_INICIO_TRAN_CASE "
sqlCase = sqlCase & "FROM CASE_CONDICAO_TRANS "

if strCase <> "0" then
	sqlCase = sqlCase & "WHERE CASE_TX_CD_CASE = '" & strCase & "'"
	sqlCase = sqlCase & " ORDER BY CASE_TX_CD_CASE"
end if

set rstCase = db_cogest.execute(sqlCase)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>	
		<STYLE type=text/css>
			BODY {
				SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
		</STYLE>			
	</head>	
	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
	
		<script language="javascript">		
			var janela;	
			
			function abrir_janela(strAcao,strCdDesenv,strDTPrevRealiz,strDTConclusao,strDTIniTrans,strMSG)
			{						
				if (janela != null)
				{					
					janela.close();					
					janela = null;					
				}				
				//alert(strAcao+"-"+strCdDesenv+"-"+strDTPrevRealiz+"-"+strDTConclusao+"-"+strDTIniTrans+"-"+strMSG);				
				janela = window.open ('inf_case_desenv.asp?CdDesenv='+strCdDesenv+'&DTPrevRealiz='+strDTPrevRealiz+'&DTConclusao='+strDTConclusao+'&DTIniTrans='+strDTIniTrans+'&msg='+strMSG,'','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=550px,height=280,top=150,left=150');
			}
		</script>
	
		<form method="POST" action="" name="frm1">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
			  <tr>
				<td width="20%" height="20">&nbsp;</td>
				<td width="44%" height="60">&nbsp;</td>
				<td width="36%" valign="top"> 
				  <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
					<tr> 
					  <td bgcolor="#330099" width="39" valign="middle" align="center"> 
						<div align="center">
						  <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Case/voltar.gif"></a></div>
					  </td>
					  <td bgcolor="#330099" width="36" valign="middle" align="center"> 
						<div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Case/avancar.gif"></a></div>
					  </td>
					  <td bgcolor="#330099" width="27" valign="middle" align="center"> 
						<div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Case/favoritos.gif"></a></div>
					  </td>
					</tr>
					<tr> 
					  <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
						<div align="center"><a href="javascript:print()"><img border="0" src="../Case/imprimir.gif"></a></div>
					  </td>
					  <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
						<div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Case/atualizar.gif"></a></div>
					  </td>
					  <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
						<div align="center"><a href="../../indexA.asp"><img src="../Case/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
					  </td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr bgcolor="#00FF99">
				<td colspan="3" height="20">
				  <table width="625" border="0" align="center">
					<tr>
					  <td width="26"></td>
					  <td width="50"></td>
					  <td width="26"></td>
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
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td height="10" colspan="2">
				  </td>
				</tr>
				<tr>
				  <td>
					<div align="center">
				    <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">&nbsp;Relatório de CASE - CONDIÇÃO - TRANSAÇÃO</font>
					</div>
				  </td>
				  <td align="right"><font face="Verdana" color="#330099" size="1">Clique sobre a imagem 
				  	<img src="../../imagens/aprova_02.gif" border="0">&nbsp;ou&nbsp;
					<img src="../../imagens/aprova_03.gif" border="0">&nbsp;para obter informações sobre os Desenvolvimentos.&nbsp</font>
				  </td>
				</tr>
				<tr>
				  <td height="10"  colspan="2">
				  </td>
				</tr>
			</table>			
			<p style="margin-top: 0; margin-bottom: 0">			
			<table border="0" width="100%">
			  <tr>
				<td width="36%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Case</font></b></td>
				<td width="35%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Condição</font></b></td>
				<td width="14%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Transações</font></b></td>
				<td width="15%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Desenvolvimentos</font></b></td>
			  </tr>
			  <%  			       
			  tem = 0          
			  do until rstCase.eof = true			  				
				  atualCase1 = rstCase("CASE_TX_CD_CASE")
				  atualCondicao2 = rstCase("CASE_NR_CD_CONDICAO")				 
				  atualTransacao3 = rstCase("TRAN_CD_TRANSACAO")						 
				  %>
				  <tr>
					<%
					if atualCase1 <> antCase1 then						
						strCase = rstCase("CASE_TX_DESC_CASE")            
					else
						strCase = ""
					end if
					
					if strCase = "" then
						cor = "white"
					else
						cor = "#CCCCCC"
					end if			
					%>
					<td width="36%" bgcolor="<%=cor%>"><font face="Verdana" size="2" color="#330099"><b><%=strCase%></b></font></td>
					<%
					if atualCondicao2 <> antCondicao2 then						 
						strCondicao = rstCase("CASE_TX_DESC_CONDICAO")       
					else
						strCondicao = ""
					end if
					
					if strCondicao = "" then
						cor = "white"
					else
						cor = "#EAEAEA"	
					end if							
					%>
					<td width="35%" bgcolor="<%=cor%>"><font face="Verdana" size="2" color="#330099"><b><%=strCondicao%></b></font></td>
					<%
					if atualTransacao3 <> antTransacao3 then						
						intCountDesenv = 0
						set rstTransacao = db_cogest.execute("SELECT TRAN_CD_TRANSACAO, TRAN_TX_DESC_TRANSACAO FROM TRANSACAO WHERE TRAN_CD_TRANSACAO ='" & rstCase("TRAN_CD_TRANSACAO") & "'")
						strTransacao = rstTransacao("TRAN_CD_TRANSACAO")
						
						rstTransacao.close
						set rstTransacao =nothing											   
					else
						if atualCondicao2 <> antCondicao2 then
							intCountDesenv = 0
							cor = "#FFFFDF"	
						else
							strTransacao = ""
						end if
						
					end if
					
					if strTransacao = "" then
						cor = "white"
					else
						cor = "#ECE9CF"	
					end if	   
					%>
					<td width="14%" bgcolor="<%=cor%>"><font face="Verdana" size="2" color="#330099"><b><%=strTransacao%></b></font></td>
					<%						
					sqlDeserv = ""							
					sqlDeserv = sqlDeserv & "SELECT DESE_CD_DESENVOLVIMENTO " 							
					sqlDeserv = sqlDeserv & "FROM TRANSACAO_DESENV "
					sqlDeserv = sqlDeserv & "WHERE TRAN_CD_TRANSACAO = '" & rstCase("TRAN_CD_TRANSACAO") & "'"
					sqlDeserv = sqlDeserv & " ORDER BY DESE_CD_DESENVOLVIMENTO "						
					'Response.write sqlDeserv
					'Response.end
					set rstDesenv = db_cogest.execute(sqlDeserv)						
																							
					do until rstDesenv.eof = true
						strNomeDesenv = rstDesenv("DESE_CD_DESENVOLVIMENTO")            
													
						'*** PARA VERIFICAR A CONDIÇÃO DOD DESENVOLVIMENTOS E DEFINIR A IMAGEM- OBS PEGA OS COM PROBLEMAS						
						sqlDeservIndic = ""							
						sqlDeservIndic = sqlDeservIndic & "SELECT DISTINCT TR_DESENV.DESE_CD_DESENVOLVIMENTO, DESENV.DESE_DT_PREVISTA_REALIZACAO, DESENV.DESE_DT_CONCLUSAO " 							
						sqlDeservIndic = sqlDeservIndic & "FROM TRANSACAO_DESENV TR_DESENV, "
						sqlDeservIndic = sqlDeservIndic & "CASE_CONDICAO_TRANS CASE_COND, "
						sqlDeservIndic = sqlDeservIndic & "DESENVOLVIMENTO DESENV "
						sqlDeservIndic = sqlDeservIndic & "WHERE CASE_COND.TRAN_CD_TRANSACAO = TR_DESENV.TRAN_CD_TRANSACAO "							
						sqlDeservIndic = sqlDeservIndic & "AND TR_DESENV.DESE_CD_DESENVOLVIMENTO = '" & rstDesenv("DESE_CD_DESENVOLVIMENTO") & "' "
						sqlDeservIndic = sqlDeservIndic & "AND TR_DESENV.DESE_CD_DESENVOLVIMENTO = DESENV.DESE_CD_DESENVOLVIMENTO "
						'sqlDeservIndic = sqlDeservIndic & "AND ((DESENV.DESE_DT_CONCLUSAO = '' OR DESENV.DESE_DT_CONCLUSAO IS NULL) "
						'sqlDeservIndic = sqlDeservIndic & "AND DESENV.DESE_DT_PREVISTA_REALIZACAO < GETDATE()) "										
						
						sqlDeservIndic = sqlDeservIndic & "AND ((DESENV.DESE_DT_CONCLUSAO = '' OR DESENV.DESE_DT_CONCLUSAO IS NULL) "
						sqlDeservIndic = sqlDeservIndic & "AND (DESENV.DESE_DT_PREVISTA_REALIZACAO < GETDATE() "
						sqlDeservIndic = sqlDeservIndic & "	OR (DESENV.DESE_DT_PREVISTA_REALIZACAO = ''	"
						sqlDeservIndic = sqlDeservIndic & "	OR DESENV.DESE_DT_PREVISTA_REALIZACAO IS NULL))) "
																	
						sqlDeservIndic = sqlDeservIndic & "AND (DESENV.DESE_DT_PREVISTA_REALIZACAO > CASE_COND.CASE_DT_INICIO_TRAN_CASE  "
						sqlDeservIndic = sqlDeservIndic & "OR CASE_COND.CASE_DT_INICIO_TRAN_CASE < GETDATE()) "
						'Response.write sqlDeservIndic & "<br><br>"						
						set rstDesenvIndic = db_cogest.execute(sqlDeservIndic)
						
						strNomeImagem = ""
						strMSG = ""
						strTitle = ""
						if not rstDesenvIndic.eof then
							strNomeImagem 	= "aprova_03.gif"																				
							strCdDesenv 	= rstDesenvIndic("DESE_CD_DESENVOLVIMENTO")
							strDTPrevRealiz	= rstDesenvIndic("DESE_DT_PREVISTA_REALIZACAO")
							strDTConclusao	= rstDesenvIndic("DESE_DT_CONCLUSAO")
							strDTIniTrans 	= rstCase("CASE_DT_INICIO_TRAN_CASE")							
							strTitle 		= "Desenvolvimento com Problemas!"
							strMSG 			= ""								
						else
							'*** PARA VERIFICAR A CONDIÇÃO DOD DESENVOLVIMENTOS E DEFINIR A IMAGEM- OBS PEGA OS SEM PROBLEMAS
							sqlDesenvIndic2 = ""							
							sqlDesenvIndic2 = sqlDesenvIndic2 & "SELECT DISTINCT TR_DESENV.DESE_CD_DESENVOLVIMENTO, DESENV.DESE_DT_PREVISTA_REALIZACAO, DESENV.DESE_DT_CONCLUSAO " 							
							sqlDesenvIndic2 = sqlDesenvIndic2 & "FROM TRANSACAO_DESENV TR_DESENV, "
							sqlDesenvIndic2 = sqlDesenvIndic2 & "CASE_CONDICAO_TRANS CASE_COND, "
							sqlDesenvIndic2 = sqlDesenvIndic2 & "DESENVOLVIMENTO DESENV "
							sqlDesenvIndic2 = sqlDesenvIndic2 & "WHERE CASE_COND.TRAN_CD_TRANSACAO = TR_DESENV.TRAN_CD_TRANSACAO "							
							sqlDesenvIndic2 = sqlDesenvIndic2 & "AND TR_DESENV.DESE_CD_DESENVOLVIMENTO = '" & rstDesenv("DESE_CD_DESENVOLVIMENTO") & "'"
							sqlDesenvIndic2 = sqlDesenvIndic2 & " AND TR_DESENV.DESE_CD_DESENVOLVIMENTO = DESENV.DESE_CD_DESENVOLVIMENTO "
							
							set rstDesenvIndic2 = db_cogest.execute(sqlDesenvIndic2)
																					
							strNomeImagem = "aprova_02.gif"
							strCdDesenv 	= strNomeDesenv
							strDTPrevRealiz	= rstDesenvIndic2("DESE_DT_PREVISTA_REALIZACAO")
							strDTConclusao	= rstDesenvIndic2("DESE_DT_CONCLUSAO")
							strDTIniTrans 	= rstCase("CASE_DT_INICIO_TRAN_CASE")							
							strMSG = "Não existe problemas com este Desenvolvimento!"
							strTitle = "Desenvolvimento - OK!"		
							
							rstDesenvIndic2.close
							set rstDesenvIndic2 = nothing					
						end if		
													
						rstDesenvIndic.close
						set rstDesenvIndic = nothing		
						
						if intCountDesenv = 0 then
						%>
								<td width="15%" bgcolor="#F4F3EA">
									<table width="100%"  border="0" cellspacing="0" cellpadding="0">
									  <tr>
										<td width="7%"></td>
										<td width="54%" align="left"><font face="Verdana" size="2" color="#330099"><b><%=strNomeDesenv%></b></font></td>
										<td width="39%">
											<a href="#"><img src="../../imagens/<%=strNomeImagem%>" border="0" title="<%=strTitle%>" onClick="javascript:abrir_janela('Esconde','<%=strCdDesenv%>','<%=strDTPrevRealiz%>','<%=strDTConclusao%>','<%=strDTIniTrans%>','<%=strMSG%>');"></a>
										</td>
									  </tr>
									</table>							
								</td>
							</tr>
						<%else%>
								<td bgcolor="#FFFFFF">
								<td bgcolor="#FFFFFF">
								<td bgcolor="#FFFFFF">
								<td width="15%" bgcolor="#F4F3EA">
									<table width="100%"  border="0" cellspacing="0" cellpadding="0">
									  <tr>
										<td width="7%"></td>
										<td width="54%" align="left"><font face="Verdana" size="2" color="#330099"><b><%=strNomeDesenv%></b></font></td>
										<td width="39%">
											<a href="#"><img src="../../imagens/<%=strNomeImagem%>" border="0" title="<%=strTitle%>" onClick="javascript:abrir_janela('Esconde','<%=strCdDesenv%>','<%=strDTPrevRealiz%>','<%=strDTConclusao%>','<%=strDTIniTrans%>','<%=strMSG%>');"></a>								
										</td>
									  </tr>
									</table>		
								</td>
							</tr>
					  <%end if%>
					  <%  
						intCountDesenv = intCountDesenv + 1						  
						rstDesenv.movenext
					loop 
			 
					rstDesenv.close
					set rstDesenv = nothing
				 
				  tem = tem + 1
				  
				  antCase1 		= rstCase("CASE_TX_CD_CASE")
				  antCondicao2 	= rstCase("CASE_NR_CD_CONDICAO")
				  antTransacao3 = rstCase("TRAN_CD_TRANSACAO")					 
							
			  	rstCase.movenext
			  loop 
			  
			  rstCase.close
			  set rstCase = nothing
			  %>			  
			</table>		
		
		<%if tem = 0 then%>
			<font face="Verdana" size="2" color="#800000">&nbsp;<b>Não foi encontrado nenhum registro para esta consulta.</b></font></b>
		<%end if%>
	  <br>				
	  </form>	
	</body>
</html>