<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

if request("pCorte") <> "0" and request("pCorte") <> "" then
	Session("Corte") = request("pCorte")
elseif request("selCorte") <> "" then
	Session("Corte") = request("selCorte")
end if
	
if trim(request("selSala")) <> "0" and trim(request("selSala")) <> "" then	
	
	strSala = trim(request("selSala"))
	
	strSQLNomeSala = ""
	strSQLNomeSala = strSQLNomeSala & "SELECT SALA_TX_NOME_SALA " 
	strSQLNomeSala = strSQLNomeSala & "FROM GRADE_SALA "
	strSQLNomeSala = strSQLNomeSala & "WHERE SALA_CD_SALA = " & strSala
	strSQLNomeSala = strSQLNomeSala & "AND CORT_CD_CORTE = " & Session("Corte")
	set rdsNomeSala = db_banco.Execute(strSQLNomeSala)
	
	if not rdsNomeSala.eof then
		strNomeSala = rdsNomeSala("SALA_TX_NOME_SALA")
	else
		strNomeSala = ""
	end if
	
	rdsNomeSala.close
	set rdsNomeSala = nothing
else
	strSala = "0"
	strNomeSala = ""
end if

strSQLSala = ""
strSQLSala = strSQLSala & "SELECT SALA_CD_SALA, SALA_TX_NOME_SALA " 
strSQLSala = strSQLSala & "FROM GRADE_SALA "
strSQLSala = strSQLSala & "WHERE CORT_CD_CORTE = " & Session("Corte")
strSQLSala = strSQLSala & " ORDER BY SALA_TX_NOME_SALA"
'Response.write strSQLSala & "<br>"
'Response.end

set rdsSala = db_banco.Execute(strSQLSala)

strSQLTurma = ""
strSQLTurma = strSQLTurma & "SELECT CORT_CD_CORTE, TURM_NR_CD_TURMA, CURS_CD_CURSO, "
strSQLTurma = strSQLTurma & "USMA_CD_USUARIO, TURM_TX_DESC_TURMA, TURM_TX_MANDANTE, "
strSQLTurma = strSQLTurma & "TURM_DT_INICIO, TURM_DT_TERMINO, TURM_HR_INICIO, "
strSQLTurma = strSQLTurma & "TURM_HR_TERMINO, TURM_NUM_QTE_PERIODO, MULT_NR_CD_ID_MULT "
strSQLTurma = strSQLTurma & "FROM GRADE_TURMA "

if strSala <> "0" then
	strSQLTurma = strSQLTurma & "WHERE SALA_CD_SALA = " & strSala 
else
	strSQLTurma = strSQLTurma & "WHERE SALA_CD_SALA = 999999"	
end if

strSQLTurma = strSQLTurma & " AND CORT_CD_CORTE = " & Session("Corte")
strSQLTurma = strSQLTurma & " ORDER BY TURM_TX_DESC_TURMA"
'Response.WRITE strSQLTurma
'Response.END
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		<style type="text/css">
		<!--
			body 
			{
				margin-left: 0px;
				margin-top: 0px;
				margin-right: 0px;
				margin-bottom: 0px;
			}
			a {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333; text-decoration: none}
			a:hover {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333;  text-decoration: underline}
		-->
		</style>	
	
		<script language="JavaScript">	
		
			function submet_pagina(strValor, strTipo)
			{					
				if (strTipo == 'Sala')
				{
					window.location.href = "inclui_altera_turma.asp?selSala="+strValor+"&selCorte="+document.frmIncAltCurso.selCorte.value;
				}								
			}
		
			function Habilita(form)
			{
			if ( form.tipo.value == 2)
				{
				form.cdassi.disabled = false
				form.cdassi.style.backgroundColor = "#FFFFFF"
				}
			else
				{
				form.cdassi.disabled = true
				form.cdassi.style.backgroundColor = "#CCCCCC"
				}
			}
			
			function confirma_Exclusao(strSala,strTurma)
			{			
				var strCorte;
			
				if (document.frmIncAltCurso.selCorte.selectedIndex == 0)
				{
					alert('Para esta operação é neccessária a escolha de um corte!');
					document.frmIncAltCurso.selCorte.focus();
					return;				
				}
				else
				{
				strCorte = document.frmIncAltCurso.selCorte.value;
				}				
						
				if(confirm("Confirma a exclusão deste Registro?"))
				{							
					document.frmIncAltCurso.action='grava_turma.asp?parAcao=E&hdSala='+strSala+'&txtCdTurma='+strTurma+'&txtCorte='+strCorte;					
					document.frmIncAltCurso.submit();	
				}
			}				
		</script>
	</head>

	<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	
		<form name="frmIncAltCurso" method="post">		
		
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
				<td colspan="3" height="31">
				  <table width="800" border="0" align="center">
					<tr>
					  <%									  
					  set rsCentro = db_banco.execute("SELECT * FROM GRADE_SALA WHERE SALA_CD_SALA=" & strSala)
					  
					  if not rsCentro.eof then
						  ct = rsCentro("CTRO_CD_CENTRO_TREINAMENTO")
						  %>
						  <td width="26" height="23"><a href="gera_grade.asp?selCT=<%=ct%>&selCorte=<%=Session("Corte")%>&pTituloRel=GRADE%20DE%20TREINAMENTO" target="_blank"><img src="../../imagens/calender.gif" width="22" height="21" border="0"></a></td>
						  <td width="283"><font color="#330099" face="Verdana" size="2"><b>Visualizar Grade de Treinamento</b></font></td>
						  <%
					  else
					  %>
					  	 <td width="54"></td>
						  <td width="31"></td>
					  <%
					  end if	
					  
					  rsCentro.close
					  set rsCentro = nothing  
					  %>
					  <td width="13">&nbsp;</td>
					  <td width="7"></td>
					  <td width="12"></td>  
						<td width="118"></td>
					  <td width="12"></td>
					  <td width="12">&nbsp;</td>
					  <td width="20"></td>
					</tr>
				  </table>				</td>
			  </tr>
			</table>
				
			<table width="90%" border="0" cellpadding="0" cellspacing="0">
			  <tr>				
				<td colspan="4" height="15"></td>
			  </tr>
			  <tr>		
			  	<td colspan="1"></td>						
				<td colspan="3" align="center">
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Turma - Grade de Treinamento</b></font></div>
				</td>
			  </tr>
			  <tr>				
				<td colspan="4" height="20"></td>
			  </tr>
			   <tr>
			     <td height="31" colspan="1"></td>
			     <td width="9%" valign="middle">
				  	<%
				 	strSQLCorte = ""
					strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
					strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
					strSQLCorte = strSQLCorte & "ORDER BY CORT_DT_DATA_CORTE DESC"
					'Response.write strSQLCorte
					'Response.end
					set rsCorte = db_banco.Execute(strSQLCorte)				  
				 	%>				   
				   <font face="Verdana" size="2" color="#330099"><b>Corte:&nbsp;</b></font>
				 </td>
			     <td width="82%" colspan="2" valign="middle">				
				   
				   <select name="selCorte" size="1" onchange="javascript:submet_pagina(this.value,'');">
					  <option value="0">== Selecione um Corte ==</option>
					  <%										
						do until rsCorte.eof=true											
							if cint(Session("Corte")) = cint(rsCorte("CORT_CD_CORTE")) then
								%>
							<option selected value="<%=rsCorte("CORT_CD_CORTE")%>"><%=rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")%></option>
					  <% 
							else							
								%>
							<option value="<%=rsCorte("CORT_CD_CORTE")%>"><%=rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")%></option>
					  <% 
							end if							
							rsCorte.movenext
						loop
						%>
					</select>		
					<%
					rsCorte.close
					set rsCorte = nothing			 
				 	%>				   	   
				</td>
		      </tr>
			
			   <tr>		
			   	<td width="9%" height="40" colspan="1"></td>		
				<td valign="middle"><font face="Verdana" size="2" color="#330099"><b>Sala:&nbsp;</b></font></td>
				<td colspan="2" valign="middle">					
					<select size="1" name="selSala" onchange="javascript:submet_pagina(this.value,'Sala');">
					  <option value="0" selected>== Selecione a Sala ==</option>
					  	<%
						do until rdsSala.eof = true
							if cint(strSala) = cint(rdsSala("SALA_CD_SALA")) then%>
					  			<option value="<%=rdsSala("SALA_CD_SALA")%>" selected><%=rdsSala("SALA_TX_NOME_SALA")%></option>
					  		<%else%>
					  			<option value="<%=rdsSala("SALA_CD_SALA")%>"><%=rdsSala("SALA_TX_NOME_SALA")%></option>
					  		<%end if						
								rdsSala.movenext
						loop
						
						rdsSala.close
						set rdsSala = nothing						
						%>
        			</select>
				</td>				
			  </tr>
			   <tr>				
				<td colspan="4" height="50"></td>
			  </tr>
			</table>
			   
			<table width="100%"  border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td height="2" bgcolor="#CCCCCC"></td>
			  </tr>
			  <tr>
				<td>&nbsp;</td>
			  </tr>
			  <tr>
				<td><div align="center"><font face="Verdana" color="#330099" size="3"><b>Turmas para Grade de Treinamento</b></font></div></td>
			  </tr>
			  <tr><td></td></tr>
			</table>
			<br>
			<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
			  <tr bgcolor="#CCCCCC"> 
			  
			  	<%if strSala <> "" and strSala <> "0" then%>			  
					<td width="5%" bgcolor="#9C9A9C" class="titcoltabela">
						&nbsp;<a href="cadastra_turma.asp?pAcao=I&pSala=<%=strSala%>"><img src="../../imagens/botao_novo_off_02.gif" alt="Incluir Turma na Grade de Treinamento" width="34" height="21" border="0"></a>
					</td>
				<%end if%>
				
				<td colspan="7">         
					<%						
					if strSala <> "" and strSala <> "0" then
						strTexto = strTexto & " Número da sala selecionada - " & strNomeSala
					else
						strTexto = strTexto & " Selecione uma Sala para montar a Grade de Treinamento"
					end if	
					%>
					<div align="center" class="campob">										
						<font face="Verdana" color="#330099" size="2"><b><%=strTexto%></b></font>
					</div>
				</td>
			  </tr>
			  <tr bgcolor="#CCCCCC">
			    <%if strSala <> "" and strSala <> "0" then%>
					<td bgcolor="#9C9A9C" class="titcoltabela">&nbsp;</td>
			    <%end if%>
				<td width="11%"><div align="center"><font face="Verdana" color="#330099" size="1"><b>Turma</b></font></div></td>
				<td width="6%"><div align="center"><font face="Verdana" color="#330099" size="1"><b>Curso</b></font></div></td>
				<td width="27%"><div align="center"><font face="Verdana" color="#330099" size="1"><b>Multiplicador</b></font></div></td>
				<td width="18%"><div align="center"><font face="Verdana" color="#330099" size="1"><b>Mandante</b></font></div></td>
				<td width="13%"><div align="center"><font face="Verdana" color="#330099" size="1"><b>Data de Início</b></font></div></td>
				<td width="13%"><div align="center"><font face="Verdana" color="#330099" size="1"><b>Data de Término</b></font></div></td>			
			  	<td width="7%"><div align="center"><font face="Verdana" color="#330099" size="1"><b>Nº de Períodos</b></font></div></td>			
			  </tr>
			  <%		
			  
			set rdsTurma = db_banco.Execute(strSQLTurma)
			if not rdsTurma.EOF then 
				  Do while not rdsTurma.EOF
					%>
					  <tr bgcolor="#E9E9E9">
						<td bgcolor="#9C9A9C">&nbsp;<a href="cadastra_turma.asp?pAcao=A&pSala=<%=strSala%>&pTurma=<%=rdsTurma("TURM_NR_CD_TURMA")%>"><img src="../../imagens/botao_abrir_off_02.gif" alt="Alterar Turma na Grade de Treinamento" width="34" height="21" border="0"></a>&nbsp;<a href="#" onClick="Javascript:confirma_Exclusao('<%=strSala%>','<%=rdsTurma("TURM_NR_CD_TURMA")%>');"><img src="../../imagens/botao_deletar_off_02.gif" alt="Excluir Turma da Grade de Treinamento" width="34" height="21" border="0"></a></td>
						<td bgcolor="#FFFFFF">
							<div align="center">
								<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rdsTurma("TURM_TX_DESC_TURMA")%></font>
							</div>
						</td>				
						<td bgcolor="#FFFFFF">				
							<div align="center">
								<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rdsTurma("CURS_CD_CURSO")%></font>
							</div>
						</td>
						
						<%																
						if trim(rdsTurma("USMA_CD_USUARIO")) <> "" then
							strMULT = rdsTurma("USMA_CD_USUARIO")
						else
							if trim(rdsTurma("MULT_NR_CD_ID_MULT")) <> "" then
							
								strSQLVeriMultTurma = ""			
								strSQLVeriMultTurma = strSQLVeriMultTurma & "SELECT DISTINCT MULT.MULT_TX_NOME_MULTIPLICADOR, MULT.MULT_NR_CD_CHAVE, MULT_TX_TIPO_MULTIPLICADOR "
								strSQLVeriMultTurma = strSQLVeriMultTurma & "FROM GRADE_TURMA TURMA, GRADE_MULTIPLICADOR MULT "
								strSQLVeriMultTurma = strSQLVeriMultTurma & "WHERE TURMA.MULT_NR_CD_ID_MULT = MULT.MULT_NR_CD_ID_MULT "
								strSQLVeriMultTurma = strSQLVeriMultTurma & "AND TURMA.MULT_NR_CD_ID_MULT = " & rdsTurma("MULT_NR_CD_ID_MULT") 
								strSQLVeriMultTurma = strSQLVeriMultTurma & " AND TURMA.CORT_CD_CORTE = " & Session("Corte") 	
								strSQLVeriMultTurma = strSQLVeriMultTurma & " AND MULT.CORT_CD_CORTE = " & Session("Corte") 	
								'Response.write strSQLVeriMultTurma & "<br>"
								'Response.end		
										
								set rsVeriMultTurma = db_banco.Execute(strSQLVeriMultTurma)		
										
								if not rsVeriMultTurma.eof then	
									blnPodeDeletar = False
									strAssociacao = "Turma"
									
									if rsVeriMultTurma("MULT_TX_TIPO_MULTIPLICADOR") <> "EXTRA" then									
										strMult = rsVeriMultTurma("MULT_TX_NOME_MULTIPLICADOR") & " - " & rsVeriMultTurma("MULT_NR_CD_CHAVE")
									else
										strMult = rsVeriMultTurma("MULT_TX_NOME_MULTIPLICADOR")
									end if
								end if
								
								rsVeriMultTurma.close
								set rsVeriMultTurma = nothing
							else						
								strMULT = "N/A"
							end if
						end if					
						%>
						<td bgcolor="#FFFFFF">
							<div align="center">
								<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strMULT%></font>
							</div>
						</td>
						
						<%						
						if trim(rdsTurma("TURM_TX_MANDANTE")) <> "" then
							strMandanteTXT = rdsTurma("TURM_TX_MANDANTE")
						else
							strMandanteTXT = "N/A"
						end if					
						%>
						<td bgcolor="#FFFFFF">
							<div align="center">
								<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strMandanteTXT%></font>
							</div>
						</td>						
						<% 
						if trim(rdsTurma("TURM_DT_INICIO")) <> "" then
							strDataIni1 = MontaDataHora(trim(rdsTurma("TURM_DT_INICIO")),2)
						else
							strDataIni1 = ""
						end if							
						if trim(rdsTurma("TURM_HR_INICIO")) <> "" then
							strHoraIni1 = " - " & MontaDataHora(trim(rdsTurma("TURM_HR_INICIO")),3)				
						else
							strHoraIni1 = ""
						end if						
						strDataIni_Hora = strDataIni1 & strHoraIni1
						%>
						<td bgcolor="#FFFFFF">
							<div align="center">
								<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strDataIni_Hora%></font>
							</div>
						</td>			
						
						<% 		
						if trim(rdsTurma("TURM_DT_TERMINO")) <> "" then
							strDataIni2 = MontaDataHora(trim(rdsTurma("TURM_DT_TERMINO")),2)
						else
							strDataIni2 = ""
						end if							
						if trim(rdsTurma("TURM_HR_TERMINO")) <> "" then
							strHoraIni2 =  " - " & MontaDataHora(trim(rdsTurma("TURM_HR_TERMINO")),3)				
						else
							strHoraIni2 = ""
						end if						
						strDataFim_Hora = strDataIni2 & strHoraIni2								
						%>
						<td bgcolor="#FFFFFF">
							<div align="center">
								<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=strDataFim_Hora%></font>
							</div>
						</td>		
						
						<td bgcolor="#FFFFFF">
							<div align="center">
								<font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=Trim(rdsTurma("TURM_NUM_QTE_PERIODO"))%></font>
							</div>
						</td>				
					  </tr>
					  <%       
			  		rdsTurma.movenext 
				 Loop 
			end if	
			
			rdsTurma.close
			set rdsTurma = nothing
			%>			
			</table>
		</form>
		<%
		
		db_banco.close
		set db_banco = nothing
		
		'db_Cogest.close
		'set db_Cogest = nothing
		%>
       	<p>&nbsp;</p>
	</body>		
	<%
	public function MontaDataHora(strData,intDataTime)

		'*** intDataTime - Indica se mostraá a data c/ hora ou apenas a data.
		'*** intDataTime = 1 (DATA E HORA)
		'*** intDataTime = 2 (DATA)
		'*** intDataTime = 3 (HORA)
		'*** intDataTime = 4 (FORMATO DE BANCO)
	
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
			MontaDataHora = strDia & "/" & strMes & "/" & year(strData) 
		elseif cint(intDataTime) = 3 then	
			MontaDataHora = strHora & ":" & strMinuto	
		elseif cint(intDataTime) = 4 then
			MontaDataHora = strMes & "/" & strDia & "/" & year(strData)
		end if
	end function
	%>
</html>