<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
strAcao = trim(Request("pAcao"))

if trim(Request("selCorte")) <> "" then
	session("Corte") = trim(Request("selCorte"))
end if

intCdCorte = session("Corte")

'intCdMultiplicador = trim(Request("selMultiplicador"))
'Response.write strAcao & "<br>"
'Response.write intCdMultiplicador & "<br>"
'Response.end

strMostraCurso = "Não"
strMostraUnid = "Não"

if strAcao = "I" then
	strNomeAcao = "Inclusão"
elseif strAcao = "A" then 
	strNomeAcao ="Alteração"
end if  

if strAcao = "A" then	
		
	if Request("selMultiplicador") <> "" and Request("selMultiplicador") <> "0" then
		strVetMult = split(Request("selMultiplicador"),"|")
	
		intCdMultiplicador = strVetMult(0)
		intTipoMultiplicador = strVetMult(1)
	end if	
		
	strSQLAltMultiplicador = ""
	strSQLAltMultiplicador = strSQLAltMultiplicador & "SELECT CORT_CD_CORTE, MULT_NR_CD_ID_MULT, MULT_TX_NOME_MULTIPLICADOR, ORME_CD_ORG_MENOR, "
	strSQLAltMultiplicador = strSQLAltMultiplicador & "MULT_TX_TIPO_MULTIPLICADOR, MULT_TX_RESTRICAO_VIAGEM, MULT_NR_TIPO_MULTIPLICADOR "
	strSQLAltMultiplicador = strSQLAltMultiplicador & "FROM GRADE_MULTIPLICADOR " 
	strSQLAltMultiplicador = strSQLAltMultiplicador & "WHERE MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	strSQLAltMultiplicador = strSQLAltMultiplicador & " AND CORT_CD_CORTE = " & intCdCorte 
	strSQLAltMultiplicador = strSQLAltMultiplicador & " AND MULT_NR_TIPO_MULTIPLICADOR = " & intTipoMultiplicador
	'Response.write strSQLAltMultiplicador
	'Response.end
	
	Set rdsAltMultiplicador = db_banco.Execute(strSQLAltMultiplicador)			
	
	if not rdsAltMultiplicador.EOF then			
	
		strNomeMultiplicador	= trim(rdsAltMultiplicador("MULT_TX_NOME_MULTIPLICADOR"))		
		'strRestrViagem 			= trim(rdsAltMultiplicador("MULT_TX_RESTRICAO_VIAGEM"))
		intTipoMultiplicador 	= rdsAltMultiplicador("MULT_NR_TIPO_MULTIPLICADOR")		
		strTipoMultiplicador 	= rdsAltMultiplicador("MULT_TX_TIPO_MULTIPLICADOR")	
		intCdDiretoria		 	= rdsAltMultiplicador("ORME_CD_ORG_MENOR")					
						
		'*** REFERENTE AO CURSO ***			
		strNomeMultiplicadorCurso = ""	
		strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & "SELECT DISTINCT MULT_CURSO.CURS_CD_CURSO "
		strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
		strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & "WHERE MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
		strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & "AND MULT_CURSO.MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
		strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & " AND MULT_CURSO.CORT_CD_CORTE = " & intCdCorte 
		strNomeMultiplicadorCurso = strNomeMultiplicadorCurso & " AND MULT.CORT_CD_CORTE = " & intCdCorte 
		'Response.write strNomeMultiplicadorCurso
		'Response.end
	
		Set rdsAltMultiplicadorCurso = db_banco.Execute(strNomeMultiplicadorCurso)	
									
		if not rdsAltMultiplicadorCurso.EOF then					
			strMostraCurso = "Sim"						
		end if		
		
		'*** REFERENTE A UNIDADE ***		
		strNomeMultiplicadorUnid= ""
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & "SELECT DISTINCT UNID_ORG_MENOR.UNID_CD_UNIDADE, UNIDADE.UNID_TX_DESC_UNIDADE "
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & "FROM GRADE_MULTIPLICADOR_ORGAO_MENOR MULT_ORG_MENOR, "
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & "GRADE_UNIDADE_ORGAO_MENOR UNID_ORG_MENOR, "
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & "GRADE_UNIDADE UNIDADE "
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & "WHERE MULT_ORG_MENOR.ORME_CD_ORG_MENOR = UNID_ORG_MENOR.ORME_CD_ORG_MENOR "
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & "AND UNIDADE.UNID_CD_UNIDADE = UNID_ORG_MENOR.UNID_CD_UNIDADE "
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & "AND MULT_ORG_MENOR.MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & " AND MULT_ORG_MENOR.CORT_CD_CORTE = " & intCdCorte 
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & " AND UNID_ORG_MENOR.CORT_CD_CORTE = " & intCdCorte 
		strNomeMultiplicadorUnid = strNomeMultiplicadorUnid & " ORDER BY UNIDADE.UNID_TX_DESC_UNIDADE"
		'Response.write strNomeMultiplicadorUnid
		'Response.end
	
		Set rdsAltMultiplicadorUnid = db_banco.Execute(strNomeMultiplicadorUnid)	
									
		if not rdsAltMultiplicadorUnid.EOF then					
			strMostraUnid = "Sim"						
		end if							
			
	end if
	
	rdsAltMultiplicador.close
	set rdsAltMultiplicador = nothing	
end if

'************ DIRETORIA ****************
strSQLDiretoria =  ""
strSQLDiretoria = strSQLDiretoria & "SELECT ORLO_CD_ORG_LOT, DIRE_TX_DESC_DIRETORIA "
strSQLDiretoria = strSQLDiretoria & "FROM GRADE_DIRETORIA "
strSQLDiretoria = strSQLDiretoria & "ORDER BY DIRE_TX_DESC_DIRETORIA "
'Response.WRITE  strSQLDiretoria & "<br><br>"
'Response.END
set rdsDiretoria = db_banco.execute(strSQLDiretoria)


'************ UNIDADE ****************
'strSQLUnidade = ""
'strSQLUnidade = strSQLUnidade & "SELECT UNIDADE.UNID_TX_DESC_UNIDADE, "
'strSQLUnidade = strSQLUnidade & "UNID_ORG_MENOR.ORME_CD_ORG_MENOR "
'strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE UNIDADE,GRADE_UNIDADE_ORGAO_MENOR UNID_ORG_MENOR "
'strSQLUnidade = strSQLUnidade & "WHERE UNIDADE.UNID_CD_UNIDADE = UNID_ORG_MENOR.UNID_CD_UNIDADE "
'strSQLUnidade = strSQLUnidade & "AND UNIDADE.CORT_CD_CORTE = " & intCdCorte 
'strSQLUnidade = strSQLUnidade & " AND UNID_ORG_MENOR.CORT_CD_CORTE = " & intCdCorte 
'strSQLUnidade = strSQLUnidade & " ORDER BY UNIDADE.UNID_TX_DESC_UNIDADE"

strSQLUnidade = ""
strSQLUnidade = strSQLUnidade & "SELECT UNID_CD_UNIDADE, UNID_TX_DESC_UNIDADE "
strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE "
strSQLUnidade = strSQLUnidade & "WHERE CORT_CD_CORTE = " & intCdCorte 

if strAcao = "A" then	
	'*** PARA RETIRAR OS CADASTRADOS DA LISTA DE SELEÇÃO
	strSQLUnidade = strSQLUnidade & " AND UNID_CD_UNIDADE NOT IN "
	strSQLUnidade = strSQLUnidade & "(SELECT DISTINCT UNID_ORG_MENOR.UNID_CD_UNIDADE "
	strSQLUnidade = strSQLUnidade & "FROM GRADE_MULTIPLICADOR_ORGAO_MENOR MULT_ORG_MENOR, "
	strSQLUnidade = strSQLUnidade & "GRADE_UNIDADE_ORGAO_MENOR UNID_ORG_MENOR, "
	strSQLUnidade = strSQLUnidade & "GRADE_UNIDADE UNIDADE "
	strSQLUnidade = strSQLUnidade & "WHERE MULT_ORG_MENOR.ORME_CD_ORG_MENOR = UNID_ORG_MENOR.ORME_CD_ORG_MENOR "
	strSQLUnidade = strSQLUnidade & "AND UNIDADE.UNID_CD_UNIDADE = UNID_ORG_MENOR.UNID_CD_UNIDADE "
	strSQLUnidade = strSQLUnidade & "AND MULT_ORG_MENOR.MULT_NR_CD_ID_MULT = " & intCdMultiplicador 
	strSQLUnidade = strSQLUnidade & " AND MULT_ORG_MENOR.CORT_CD_CORTE = " & intCdCorte 
	strSQLUnidade = strSQLUnidade & " AND UNID_ORG_MENOR.CORT_CD_CORTE = " & intCdCorte & ") "	
end if

strSQLUnidade = strSQLUnidade & " ORDER BY UNID_TX_DESC_UNIDADE"
'Response.WRITE  strSQLUnidade & "<br><br>"
'Response.END
set rsUnidade = db_banco.execute(strSQLUnidade)

'************ CURSO ****************
strSQLCurso = ""
strSQLCurso = strSQLCurso & "SELECT CURS_CD_CURSO "
strSQLCurso = strSQLCurso & "FROM GRADE_CURSO " 
'strSQLCurso = strSQLCurso & "WHERE CURS_TX_CENTRALIZADO <> 'DESCENTRALIZADO' "
'strSQLCurso = strSQLCurso & "WHERE (CURS_TX_CENTRALIZADO <> 'DESCENTRALIZADO' "
'strSQLCurso = strSQLCurso & "OR CURS_TX_CENTRALIZADO IS NOT NULL "
'strSQLCurso = strSQLCurso & "OR CURS_TX_CENTRALIZADO <> '') "
strSQLCurso = strSQLCurso & "WHERE CORT_CD_CORTE = " & intCdCorte

if strAcao = "A" then	
	'*** PARA RETIRAR OS CADASTRADOS DA LISTA DE SELEÇÃO
	strSQLCurso = strSQLCurso & " AND CURS_CD_CURSO NOT IN "
	strSQLCurso = strSQLCurso & "(SELECT DISTINCT MULT_CURSO.CURS_CD_CURSO "
	strSQLCurso = strSQLCurso & "FROM GRADE_MULTIPLICADOR MULT, GRADE_MULTIPLICADOR_CURSO MULT_CURSO "
	strSQLCurso = strSQLCurso & "WHERE MULT.MULT_NR_CD_ID_MULT = MULT_CURSO.MULT_NR_CD_ID_MULT "
	strSQLCurso = strSQLCurso & "AND MULT_CURSO.MULT_NR_CD_ID_MULT = '" & intCdMultiplicador & "' "
	strSQLCurso = strSQLCurso & "AND MULT.CORT_CD_CORTE = " & intCdCorte 
	strSQLCurso = strSQLCurso & " AND MULT_CURSO.CORT_CD_CORTE = " & intCdCorte & ") "	
end if

strSQLCurso = strSQLCurso & " ORDER BY CURS_CD_CURSO"
'Response.write strSQLCurso & "<br><br><br>"
'Response.end
set rsCurso = db_banco.Execute(strSQLCurso)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		<script language="javascript" src="../js/troca_lista.js"></script>
		
		<script language="javascript">
					
			function Confirma()
			{				
				var strValorViagem;
			
				if(document.frmCadMultiplicador.txtCdMultiplicador.value == "")
				{
					alert("É necessário o preenchimento do campo CÓDIGO DO MULTIPLICADOR!");
					document.frmCadMultiplicador.txtCdMultiplicador.focus();
					return;
				}					
				
				if(document.frmCadMultiplicador.selUnidade_Selecionado.options.length == 0)
				{
					alert("Selecione pelo menos uma UNIDADE!");
					document.frmCadMultiplicador.selUnidade_Selecionado.focus();
					return;
				}													
											
				if(document.frmCadMultiplicador.selCurso_Selecionado.options.length == 0)
				{
					alert("Selecione pelo menos um CURSO!");
					document.frmCadMultiplicador.selCurso_Selecionado.focus();
					return;
				}						
																		
				//*** Monta uma string com os UNIDADES Selecionadas, separados por vírgula
				carrega_txt(document.frmCadMultiplicador.selUnidade_Selecionado,'Unidade')	
														
				//*** Monta uma string com os CURSOS Selecionados, separados por vírgula
				carrega_txt(document.frmCadMultiplicador.selCurso_Selecionado,'Curso')										
														
				//document.frmCadMultiplicador.action="grava_multiplicador.asp?pRestrViagem="+strValorViagem;		
				document.frmCadMultiplicador.action="grava_multiplicador.asp";	
				document.frmCadMultiplicador.submit();			
			}		
			
			function carrega_txt(fbox,strTipo) 
			{
				if (strTipo == 'Unidade')
				{
					document.frmCadMultiplicador.txtUnid_Selecionados.value = '';
					for(var i=0; i<fbox.options.length; i++) 
					{
						if (i == 0)
						{
							document.frmCadMultiplicador.txtUnid_Selecionados.value = fbox.options[i].value;
						}
						else
						{					
							document.frmCadMultiplicador.txtUnid_Selecionados.value = document.frmCadMultiplicador.txtUnid_Selecionados.value + "," + fbox.options[i].value;
						}
					}
				}
				
				if (strTipo == 'Curso')
				{
					document.frmCadMultiplicador.txtCurso_Selecionados.value = '';
					for(var i=0; i<fbox.options.length; i++) 
					{
						if (i == 0)
						{
							document.frmCadMultiplicador.txtCurso_Selecionados.value = fbox.options[i].value;
						}
						else
						{					
							document.frmCadMultiplicador.txtCurso_Selecionados.value = document.frmCadMultiplicador.txtCurso_Selecionados.value + "," + fbox.options[i].value;
						}
					}

				}				
			}	
			
			function submet_pagina(strValor, strTipo)
			{				
				window.location.href = "cadastra_multiplicador.asp?pAcao="+document.frmCadMultiplicador.parAcao.value+"&selUnidade=0&selCurso=0&selCorte="+strValor;											
			}
				
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmCadMultiplicador">		  
			 
			<input type="hidden" value="<%=strAcao%>" name="parAcao"> 	
			<input type="hidden" name="txtUnid_Selecionados">	
			<input type="hidden" name="txtCurso_Selecionados">	
			<input type="hidden" name="pintTipoMult" value="<%=intTipoMultiplicador%>">
			<input type="hidden" name="pstrTipoMult" value="<%=strTipoMultiplicador%>">				
									
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
						<td width="24"><a href="javascript:Confirma();"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
					  <td width="46"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
					  <td width="21">&nbsp;</td>
					  <td width="177"></td>
						<td width="30"></td>  
						<td width="234"></td>
					    <td width="9"></td>
					  <td width="8">&nbsp;</td>
					  <td width="38"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
					
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
				  <td height="10">
				  </td>
				</tr>
				<tr>
				  <td>
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Multiplicador Extra - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="695" height="124">			  	
			  	<tr>
			  	  <td height="27"></td>			  	 
			  	  <td height="27" valign="middle" align="left" colspan="2"><font face="Verdana" color="#330099" size="2"><b>Operação:</b>&nbsp;&nbsp;<%=strNomeAcao%></font></td>
		  	    </tr>
			  	<tr>
			  	  <td height="7"></td>
			  	  <td height="7" valign="middle" align="left"></td>
			  	  <td height="7" valign="middle" align="left"></td>
		  	    </tr>
								
				 <tr>
					 <td height="26" colspan="1"></td>
					 <td width="207" valign="middle">						
					   <font face="Verdana" size="2" color="#330099"><b>Corte:&nbsp;</b></font>
					 </td>
					 <td width="300" colspan="2" valign="middle">				
					   
					 <%if strAcao = "A" then
					 
					 	strSQLCorte = ""
						strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
						strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
						strSQLCorte = strSQLCorte & "WHERE CORT_CD_CORTE = " & Session("Corte")
						'Response.write strSQLCorte
						'Response.end
						set rsCorte = db_banco.Execute(strSQLCorte)		
						
						strNomeCorte = rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")
					 	%>						
						<input type="hidden" name="selCorte" value="<%=cint(Session("Corte"))%>">	
						<font face="Verdana" size="2" color="#330099"><%=Ucase(strNomeCorte)%></font>						
					 <%
					 else						 	
						strSQLCorte = ""
						strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
						strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
						strSQLCorte = strSQLCorte & "ORDER BY CORT_DT_DATA_CORTE DESC"
						'Response.write strSQLCorte
						'Response.end
						set rsCorte = db_banco.Execute(strSQLCorte)									   
					 	%>					 
					   <select name="selCorte" size="1" onchange="javascript:submet_pagina(this.value,'');">							
							<option value="0">== Selecione um Corte ==</option>
							<%										
							do until rsCorte.eof=true											
								if cint(Session("Corte")) = cint(rsCorte("CORT_CD_CORTE")) then
									%>
									<option value="<%=rsCorte("CORT_CD_CORTE")%>" selected><%=rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")%></option>
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
					end if	
					
					rsCorte.close
					set rsCorte = nothing		
					%>				   	   
					</td>
			    </tr>
				
				<!--
				<tr>
				  <td height="31"></td>
				  <td height="31" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Diretoria:</b></font></td>
				  <td height="31" valign="middle" align="left">				 
				  <select size="1" name="selDiretoria">
					<option value="0">== Selecione a Diretoria ==</option>
					<%
						'do until rdsDiretoria.eof = true
							  'if cint(intCdDiretoria) = cint(rdsDiretoria("ORLO_CD_ORG_LOT")) then%>
								<option value="<%'=rdsDiretoria("ORLO_CD_ORG_LOT")%>" selected><%'=rdsDiretoria("DIRE_TX_DESC_DIRETORIA")%></option>
							<%'else%>
								<option value="<%'=rdsDiretoria("ORLO_CD_ORG_LOT")%>"><%'=rdsDiretoria("DIRE_TX_DESC_DIRETORIA")%></option>
							<%'end if						
							'rdsDiretoria.movenext
						'loop
						
						'rdsDiretoria.close
						'set rdsDiretoria = nothing
						%>
				  </select>				  </td>
				</tr>
				-->
				
				<tr> 
				  <td width="174" height="26"></td>
				  <td width="207" height="26" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Código do Multiplicador:</b></font></td>
				  <td height="26" valign="middle" align="left" width="300">					
			 	  	<%if strAcao = "A" then%>
						<input type="hidden" name="selMultiplicador" value="<%=Request("selMultiplicador")%>">
						<input type="text" name="txtCdMultiplicador" value="<%=Ucase(strNomeMultiplicador)%>" size="15" maxlength="10">		
					<%else%>
						<input type="text" name="txtCdMultiplicador" value="<%=Ucase(intCdMultiplicador)%>" size="15" maxlength="10">	
				  <%end if%>				  
				  </td>
				</tr> 
		  </table>		  
		  
		    	<table width="856" height="175" border="0">
				<tr>
					<td width="174" height="55" rowspan="5"></td>
					<td colspan="3"><font face="Verdana" size="2" color="#330099"><b>Unidade:</b></font></td>
				</tr>
				<tr> 
				  <td width="330" height="55" rowspan="5" align="center" valign="middle">        
				  <p align="left">
				  	<font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Dispon&iacute;veis</font></p>
					<p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">
					  <select name="selUnidade" size="5" multiple>
						<%
						do until rsUnidade.eof=true
						%>
							<option value="<%=rsUnidade("UNID_CD_UNIDADE")%>"><%=rsUnidade("UNID_TX_DESC_UNIDADE")%></option>
							<%
							rsUnidade.movenext
						loop
						%>
					</select>
					</font>
					</p>
				  </td>
				  <td width="45" height="32" align="center" valign="middle"><div align="left"></div></td>
				  <td width="289" rowspan="5" align="center" valign="middle">
				  <p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Selecionadas</font></p>
					<p align="left">					
						<select name="selUnidade_Selecionado" size="5" multiple>
						<%
						if strMostraUnid = "Sim" then										
							do while not rdsAltMultiplicadorUnid.eof					
								%>
								<option value="<%=rdsAltMultiplicadorUnid("UNID_CD_UNIDADE")%>"><%=rdsAltMultiplicadorUnid("UNID_TX_DESC_UNIDADE")%></option>			
								<%								
								rdsAltMultiplicadorUnid.movenext
							loop								
				
							rdsAltMultiplicadorUnid.close
							set rdsAltMultiplicadorUnid = nothing			
						end if		
						%>					 
						</select>
					</p></td>
				</tr>
				<tr>
				  <td height="53" align="center" valign="middle"><div align="center"><img src="../../imagens/continua_F01.gif" width="24" height="24" onClick="move(document.frmCadMultiplicador.selUnidade,document.frmCadMultiplicador.selUnidade_Selecionado,1)"></div></td>
				</tr>
				<tr>
				  <td height="34" align="center" valign="middle"><div align="center"><img src="../../imagens/continua2_F01.gif" width="24" height="24" onClick="move(document.frmCadMultiplicador.selUnidade_Selecionado,document.frmCadMultiplicador.selUnidade,1)"></div></td>
				</tr>
				<tr>
				  <td height="26" align="center" valign="middle"></td>
				</tr>				
		  </table>
		  
		  	<table width="693" height="175" border="0">
				<tr>
					<td width="175" height="55" rowspan="5"></td>
					<td colspan="3"><font face="Verdana" size="2" color="#330099"><b>Curso:</b></font></td>
				</tr>
				<tr> 
				  <td width="172" height="55" rowspan="5" align="center" valign="middle">        
				  <p align="left">
				  	<font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Dispon&iacute;veis</font></p>
					<p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">
					  <select name="selCurso" size="5" multiple>
						<%
						do until rsCurso.eof=true
						%>
							<option value="<%=rsCurso("CURS_CD_CURSO")%>"><%=rsCurso("CURS_CD_CURSO")%></option>
							<%
							rsCurso.movenext
						loop
						%>
					</select>
					</font>
					</p>
				  </td>
				  <td width="36" height="32" align="center" valign="middle"><div align="left"></div></td>
				  <td width="292" rowspan="5" align="center" valign="middle">
				  <p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Selecionados</font></p>
					<p align="left">					
						<select name="selCurso_Selecionado" size="5" multiple>
						<%
						if strMostraCurso = "Sim" then										
							do while not rdsAltMultiplicadorCurso.eof					
								%>
								<option value="<%=rdsAltMultiplicadorCurso("CURS_CD_CURSO")%>"><%=rdsAltMultiplicadorCurso("CURS_CD_CURSO")%></option>			
								<%								
								rdsAltMultiplicadorCurso.movenext
							loop								
				
							rdsAltMultiplicadorCurso.close
							set rdsAltMultiplicadorCurso = nothing			
						end if		
						%>					 
						</select>
					</p></td>
				</tr>
				<tr>
				  <td height="53" align="center" valign="middle"><div align="center"><img src="../../imagens/continua_F01.gif" width="24" height="24" onClick="move(document.frmCadMultiplicador.selCurso,document.frmCadMultiplicador.selCurso_Selecionado,1)"></div></td>
				</tr>
				<tr>
				  <td height="34" align="center" valign="middle"><div align="center"><img src="../../imagens/continua2_F01.gif" width="24" height="24" onClick="move(document.frmCadMultiplicador.selCurso_Selecionado,document.frmCadMultiplicador.selCurso,1)"></div></td>
				</tr>
				<tr>
				  <td height="26" align="center" valign="middle"></td>
				</tr>				
		  </table>
		 		  
	</form>
	</body>
	<%		
	rsUnidade.close
	set rsUnidade = nothing
	
	rsCurso.close
	set rsCurso = nothing
	
	db_banco.close
	set db_banco = nothing
	%>

</html>
