<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
strAcao = trim(Request("pAcao"))
intCdFeriado = trim(Request("selFeriado"))
'Response.write strAcao & "<br>"

if Request("selCorte") <> "" then
	Session("Corte") = Request("selCorte")
end if	

if Request("rdFeriNacional") <> "" then
	strFeriNacio = Request("rdFeriNacional")
else
	strFeriNacio = "0"
end if

if Request("txtNomeFeriado") <> "" then
	strNomeFeriado = Request("txtNomeFeriado")
end if

if Request("txtDtFeriado") <> "" then
	strDtFeriado = Request("txtDtFeriado")
end if

'if Request("selCT") <> "" then
	'strCT = Request("selCT")
'else
  	'strCT = "0"
'end if

if Request("selSala_Selecionado") <> "" then
	strSala_Selecionado = Request("selSala_Selecionado")
else
  	strSala_Selecionado = "0"
end if

'Response.write trim(Request("txtMostraSala")) & "<br>"
if trim(Request("txtMostraSala")) = "S" then
	strMostraSala = "Sim"
else
	strMostraSala = "Não"
end if

'strMostraListSala = "Não"
'Response.write strMostraSala

if strAcao = "I" then
	strNomeAcao = "Inclusão"
elseif strAcao = "A" then 
	strNomeAcao ="Alteração"
end if  

'*** REFERENTE À SALA ***
strSQLSala = ""
strSQLSala = strSQLSala & "SELECT SALA.SALA_CD_SALA, SALA.SALA_TX_NOME_SALA, "
strSQLSala = strSQLSala & "SALA.CTRO_CD_CENTRO_TREINAMENTO, CT.CTRO_TX_NOME_CENTRO_TREINAMENTO "
strSQLSala = strSQLSala & "FROM GRADE_SALA SALA, GRADE_CENTRO_TREINAMENTO CT " 
strSQLSala = strSQLSala & "WHERE SALA.CTRO_CD_CENTRO_TREINAMENTO = CT.CTRO_CD_CENTRO_TREINAMENTO "
strSQLSala = strSQLSala & "AND CT.CORT_CD_CORTE = " & Session("Corte")
strSQLSala = strSQLSala & " AND SALA.CORT_CD_CORTE = " & Session("Corte")


if strAcao = "A" then	
	'*** PARA RETIRAR OS CADASTRADOS DA LISTA DE SELEÇÃO
	strSQLSala = strSQLSala & " AND SALA.SALA_CD_SALA NOT IN "
	strSQLSala = strSQLSala & "(SELECT SALA_FERIADO.SALA_CD_SALA "
	strSQLSala = strSQLSala & "FROM GRADE_FERIADO_SALA SALA_FERIADO, GRADE_SALA SALA  "
	strSQLSala = strSQLSala & "WHERE SALA.SALA_CD_SALA = SALA_FERIADO.SALA_CD_SALA "
	
	strSQLSala = strSQLSala & "AND SALA.CORT_CD_CORTE = " & Session("Corte")
	strSQLSala = strSQLSala & " AND SALA_FERIADO.CORT_CD_CORTE = " & Session("Corte")
	
	strSQLSala = strSQLSala & " AND SALA_FERIADO.FERI_CD_FERIADO = " & intCdFeriado & ")"		
end if

strSQLSala = strSQLSala & " ORDER BY SALA_TX_NOME_SALA"
'Response.write strSQLSala
'Response.end
set rsSala = db_banco.Execute(strSQLSala)

if strAcao = "A" then	
	
	strCdFeriado = trim(Request("selFeriado"))
	
	strSQLAltFeriado = ""
	strSQLAltFeriado = strSQLAltFeriado & "SELECT FERI_CD_FERIADO, FERI_TX_NOME_FERIADO, FERI_DT_DATA_FERIADO, FERI_TX_TIPO_FERIADO "
	strSQLAltFeriado = strSQLAltFeriado & "FROM GRADE_FERIADO " 
	strSQLAltFeriado = strSQLAltFeriado & "WHERE FERI_CD_FERIADO = "& strCdFeriado
	'Response.write strSQLAltFeriado
	'Response.end
	
	Set rdsAltFeriado = db_banco.Execute(strSQLAltFeriado)			
	
	if not rdsAltFeriado.EOF then	
		
		if trim(Request("pMantemCampo")) <> "S" then			
			strNomeFeriado	= rdsAltFeriado("FERI_TX_NOME_FERIADO")				
			strDtFeriado = trim(rdsAltFeriado("FERI_DT_DATA_FERIADO"))
			strFeriNacio = rdsAltFeriado("FERI_TX_TIPO_FERIADO")
		end if
					
		'*** FERIADO MUNICIPAL
		if strFeriNacio = 1 then								
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "SELECT SALA_FERIADO.SALA_CD_SALA, SALA.SALA_TX_NOME_SALA "
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "FROM GRADE_FERIADO_SALA SALA_FERIADO, GRADE_SALA SALA "
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "WHERE SALA.SALA_CD_SALA = SALA_FERIADO.SALA_CD_SALA "
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & "AND SALA_FERIADO.FERI_CD_FERIADO = " & intCdFeriado
			
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & " AND SALA.CORT_CD_CORTE = " & Session("Corte")
			strSQLIncFeriadoSala = strSQLIncFeriadoSala & " AND SALA_FERIADO.CORT_CD_CORTE = " & Session("Corte")
	
			'Response.write strSQLIncFeriadoSala
			'Response.end
		
			Set rdsAltFeriadoSala = db_banco.Execute(strSQLIncFeriadoSala)	
										
			'if not rdsAltFeriadoSala.EOF then			
				strMostraSala = "Sim"					
				'strMostraListSala = "Sim"						
			'end if				
		end if		
	end if
	
	rdsAltFeriado.close
	set rdsAltFeriado = nothing	
end if
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		<script language="javascript" src="../js/troca_lista.js"></script>
		
		<script language="javascript">
					
			function Confirma()
			{				
				if(document.frmCadFeriado.txtNomeFeriado.value == "")
				{
					alert("É necessário o preenchimento do campo DESCRIÇÃO DO FERIADO!");
					document.frmCadFeriado.txtNomeFeriado.focus();
					return;
				}					
							
				if(document.frmCadFeriado.txtDtFeriado.value == "")
				{
					alert("É necessário o preenchimento do campo DATA DO FERIADO!");
					document.frmCadFeriado.txtDtFeriado.focus();
					return;
				}						
								
				if (document.frmCadFeriado.rdFeriNacional(1).checked == true)
				{					
					if (document.frmCadFeriado.selSala_Selecionado.options.length == 0)
					{
						alert("Para a opção FERIADO MUNICIPAL,é necessária a seleção de uma SALA!");
						document.frmCadFeriado.selSala.focus();
						return;
					}
					
					//*** Monta uma string com os CTs Selecionados, separados por vírgula
					carrega_txt(document.frmCadFeriado.selSala_Selecionado)
				}
				else
				{
					document.frmCadFeriado.txtSalas_Selecionadas.value = '';
				}
				
				if (document.frmCadFeriado.selCorte.selectedIndex == 0)
				{
					alert("dddd")
					document.frmCadFeriado.selCorte.focus();
					return;
				}
														
				document.frmCadFeriado.action="grava_feriado.asp";				
				document.frmCadFeriado.submit();			
			}			
			
			function carrega_txt(fbox) 
			{
				document.frmCadFeriado.txtSalas_Selecionadas.value = '';
				for(var i=0; i<fbox.options.length; i++) 
				{
					if (i == 0)
					{
						document.frmCadFeriado.txtSalas_Selecionadas.value = fbox.options[i].value;
					}
					else
					{					
						document.frmCadFeriado.txtSalas_Selecionadas.value = document.frmCadFeriado.txtSalas_Selecionadas.value + "," + fbox.options[i].value;
					}
				}
			}

			function MostraEscondeCT(strAcao)
			{											
				if (strAcao == 'Esconde')
				{		
					document.frmCadFeriado.selCorte.disabled = true;
					//document.frmCadFeriado.selCT.disabled = true;					
					document.frmCadFeriado.selSala.disabled = true;	
					document.frmCadFeriado.selSala_Selecionado.disabled = true;								
				}
				
				if (strAcao == 'Mostra')
				{	
					document.frmCadFeriado.selCorte.disabled = false;	
					//document.frmCadFeriado.selCT.disabled = false;				   						
					document.frmCadFeriado.selSala.disabled = false;
					document.frmCadFeriado.selSala_Selecionado.disabled = false;					
				}				
			}
			
			function Verifica_Dif_Numeros(strValor,strNome)	
			{		
				var intTamanho = strValor.length;
							
				if (intTamanho == 2)
				{					
					if (isNaN(strValor.substring(0,2)) || isNaN(strValor.substring(4,5)))
					{						
						alert("O contéudo do campo Data do Feriado deve ser preenchido apenas com nº!");
						document.frmCadFeriado.txtDtFeriado.value = '';
						document.frmCadFeriado.txtDtFeriado.focus();
						return;						
					}
					
					document.frmCadFeriado.txtDtFeriado.value = document.frmCadFeriado.txtDtFeriado.value + '/';
				}
				else
				{																	
					if (isNaN(strValor.substring(0,2)) || isNaN(strValor.substring(3,5)))
					{												
						alert("O contéudo do campo Data do Feriado deve ser preenchido apenas com nº!");
						document.frmCadFeriado.txtDtFeriado.value = '';
						document.frmCadFeriado.txtDtFeriado.focus();
						return;						
					}
				}
				
				if (intTamanho == 5)
				{						
					var intDia = strValor.substring(0,2);
					var intMes = strValor.substring(3,5);
					
					if ( ((intDia == 00) || (intDia > 31)) || ((intMes == 00) || (intMes > 12)) ) 
					{
						alert('A data ' + intDia + '/' + intMes + ' é uma data inválida!');
						document.frmCadFeriado.txtDtFeriado.value = '';
						document.frmCadFeriado.txtDtFeriado.focus();
						return;	
					}
				}
			}	
			
			function submet_pagina(strValor, strTipo)
			{		
				var strAcao = document.frmCadFeriado.parAcao.value;			
				var strDtFeriado = document.frmCadFeriado.txtDtFeriado.value;
				var strNomeFeriado = document.frmCadFeriado.txtNomeFeriado.value;
				var intFerNac = 0;				
				
				document.frmCadFeriado.txtSalas_Selecionadas.value = '';		
				//document.frmCadFeriado.selSala_Selecionado.value = '';	
				
				if (document.frmCadFeriado.rdFeriNacional(1).checked == true)
				{
					intFerNac = 1;								
				}
						
				if (strTipo == 'Corte')
				{	
					//document.frmCadFeriado.action = "cadastra_feriado.asp?selSala=0&selSala_Selecionado=0&selCorte="+strValor+"&selCT=0&rdFeriNacional="+intFerNac+"&txtDtFeriado="+strDtFeriado+"&txtNomeFeriado="+strNomeFeriado+"&pAcao="+strAcao;
					document.frmCadFeriado.action = "cadastra_feriado.asp?selSala=0&selSala_Selecionado=0&selCorte="+strValor+"&rdFeriNacional="+intFerNac+"&txtDtFeriado="+strDtFeriado+"&txtNomeFeriado="+strNomeFeriado+"&pAcao="+strAcao+"&pMantemCampo=S";
				}
				
				//if (strTipo == 'CT')
				//{					
					//var strCorte = document.frmCadFeriado.selCorte.value;		
					//document.frmCadFeriado.action = "cadastra_feriado.asp?selSala=0&selSala_Selecionado=0&selCorte="+strCorte+"&selCT="+strValor+"&rdFeriNacional="+intFerNac+"&txtDtFeriado="+strDtFeriado+"&txtNomeFeriado="+strNomeFeriado+"&pAcao="+strAcao;
				//}			
				document.frmCadFeriado.submit();						
			}			
				
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmCadFeriado">		  
			 
			<input type="hidden" value="<%=strAcao%>" name="parAcao"> 	
			<input type="hidden" value="<%=intCdFeriado%>" name="txtCFeriado">		
			<input type="hidden" value="<%=intCdFeriado%>" name="selFeriado">
			<input type="hidden" name="txtSalas_Selecionadas">
			
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
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Feriado - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="884" height="255">
			  
			  	<tr>
			  	  <td height="20"></td>
			  	  <td height="20" valign="middle" align="center" colspan="2">
				  	<%if strGrava = "GravaTurma" then%>
				  	<font face="Verdana" color="#FE5A31" size="2"><b><%=strMSG%></b></font>
					<%end if%>
				  </td>			  	  
		  	    </tr>
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
				  <td width="166" height="42"></td>
				  <td width="183" height="42" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Descrição do Feriado:</b></font></td>
				  <td height="42" valign="middle" align="left" width="521">				  	
					 <input type="hidden" name="txtCdFeriado" value="<%=strCdFeriado%>">					
				 	 <input type="text" name="txtNomeFeriado" maxlength="50" size="50" value="<%=strNomeFeriado%>">				  
				  </td>
				</tr> 
								
				<tr> 
				  <td width="166" height="34"></td>
				  <td width="183" height="34" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Data do Feriado:</b></font></td>
				  <td height="34" valign="middle" align="left" width="521"> 
					<input type="text" name="txtDtFeriado" maxlength="5" size="10" value="<%=strDtFeriado%>" onKeyUp="javascript:Verifica_Dif_Numeros(this.value,this.name);" onChange="javascript:Verifica_Capacidade(this.value);">			     
				 </td>
				</tr>
												
				<tr> 
				  <td width="166" height="28"></td>
				  <td width="183" height="28" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Feriado Nacional:</b></font></td>
				  <td height="28" valign="middle" align="left" width="521">					
					<%if strFeriNacio = "0" then%>
						<input name="rdFeriNacional" type="radio" value="0" checked onClick="MostraEscondeCT('Esconde');">
					<%else%>
						<input name="rdFeriNacional" type="radio" value="0" onClick="MostraEscondeCT('Esconde');">	
					<%end if%>					   
				  </td>
				</tr>
				
				<tr> 
				  <td width="166" height="36"></td>
				  <td width="183" height="36" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Feriado Municipal:</b></font></td>
				  <td height="36" valign="middle" align="left" width="521">					
					<%if strFeriNacio = "1" then%>
						<input name="rdFeriNacional" type="radio" value="1" checked onClick="MostraEscondeCT('Mostra');"> 	
					<%else%>
						<input name="rdFeriNacional" type="radio" value="1" onClick="MostraEscondeCT('Mostra');"> 	
					<%end if%>							   
				  </td>
				</tr>			
				
				
				
				<tr>
					 <td height="43" colspan="1"></td>
					 <td width="183" valign="middle">						
					   <font face="Verdana" size="2" color="#330099"><b>Corte:</b></font>
					 </td>
					 <td width="521" colspan="2" valign="middle">				
					   
					 <%'if strAcao = "A" then
					 
					 	'strSQLCorte = ""
						'strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
						'strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
						'strSQLCorte = strSQLCorte & "WHERE CORT_CD_CORTE = " & Session("Corte")
						''Response.write strSQLCorte
						''Response.end
						'set rsCorte = db_banco.Execute(strSQLCorte)		
						
						'strNomeCorte = rsCorte("CORT_TX_DESC_CORTE") & " - " & rsCorte("CORT_DT_DATA_CORTE")
					 	%>						
						<!--<input type="hidden" name="selCorte" value="<%'=cint(Session("Corte"))%>">	
						<font face="Verdana" size="2" color="#330099"><%'=Ucase(strNomeCorte)%></font>	-->					
					 <%
					' else						 	
						strSQLCorte = ""
						strSQLCorte = strSQLCorte & "SELECT CORT_CD_CORTE, CORT_TX_DESC_CORTE, CORT_DT_DATA_CORTE "
						strSQLCorte = strSQLCorte & "FROM GRADE_CORTE " 
						strSQLCorte = strSQLCorte & "ORDER BY CORT_DT_DATA_CORTE DESC"
						'Response.write strSQLCorte
						'Response.end
						set rsCorte = db_banco.Execute(strSQLCorte)									   
					 	%>					 
					   <select name="selCorte" size="1" onchange="javascript:submet_pagina(this.value,'Corte');">							
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
					'end if	
					
					rsCorte.close
					set rsCorte = nothing		
					%>				   	   
					</td>
			    </tr>
				
				<!--
				<tr>
					 <td height="43" colspan="1"></td>
					 <td width="207" valign="middle">						
					   <font face="Verdana" size="2" color="#330099"><b>Centro de Treinamento:&nbsp;</b></font>
					 </td>
					 <td width="300" colspan="2" valign="middle">					   	 
					   <select name="selCT" size="1" onchange="javascript:submet_pagina(this.value,'CT');">							
							<option value="0">== Selecione um Centro de Treinamento ==</option>
							<%										
							'do until rsCT.eof=true											
								'if cint(strCT) = cint(rsCT("CTRO_CD_CENTRO_TREINAMENTO")) then
									%>
									<option value="<%'=rsCT("CTRO_CD_CENTRO_TREINAMENTO")%>" selected><%'=rsCT("CTRO_TX_NOME_CENTRO_TREINAMENTO")%></option>
									<% 
								'else							
									%>
									<option value="<%'=rsCT("CTRO_CD_CENTRO_TREINAMENTO")%>"><%'=rsCT("CTRO_TX_NOME_CENTRO_TREINAMENTO")%></option>
									<% 
								'end if							
								'rsCT.movenext
							'loop
							%>
					   </select>		
						<%					
						'rsCT.close
						'set rsCT = nothing		
						%>				   	   
					</td>
			    </tr>
				-->
		  </table>		  
		  
		  	<table width="885" height="175" border="0">
				<tr>
					<td width="167" height="55" rowspan="5"></td>
					<td colspan="3"><font face="Verdana" size="2" color="#330099"><b>Sala:</b></font></td>
				</tr>
				<tr> 
				  <td width="394" height="55" rowspan="5" align="center" valign="middle">        
				  <p align="left">
				  	<font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Dispon&iacute;veis:</font></p>
					<p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                    <select name="selSala" size="5" multiple>
                      <%
						do until rsSala.eof=true
						%>
                      <option value="<%=rsSala("SALA_CD_SALA")%>"><%=rsSala("SALA_TX_NOME_SALA") & " - " & rsSala("CTRO_TX_NOME_CENTRO_TREINAMENTO")%></option>
                      <%
							rsSala.movenext
						loop
						%>
                    </select>
				</font>
					</p>
				  </td>
				  <td width="50" height="32" align="center" valign="middle"><div align="left"></div></td>
				  <td width="256" rowspan="5" align="center" valign="middle">
				  <p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Selecionadas:</font></p>
					<p align="left">					
						<select name="selSala_Selecionado" size="5" multiple>
						<%
						if strMostraSala = "Sim" then										
							do while not rdsAltFeriadoSala.eof					
								%>
								<option value="<%=rdsAltFeriadoSala("SALA_CD_SALA")%>"><%=rdsAltFeriadoSala("SALA_TX_NOME_SALA")%></option>			
								<%								
								rdsAltFeriadoSala.movenext
							loop									
				
							rdsAltFeriadoSala.close
							set rdsAltFeriadoSala = nothing		
						end if		
						%>					 
						</select>
					</p></td>
				</tr>
				
				<tr>
				  <td height="53" align="center" valign="middle"><div align="center"><img src="../../imagens/continua_F01.gif" width="24" height="24" onClick="move(document.frmCadFeriado.selSala,document.frmCadFeriado.selSala_Selecionado,1)"></div></td>
				</tr>
				<tr>
				  <td height="34" align="center" valign="middle"><div align="center"><img src="../../imagens/continua2_F01.gif" width="24" height="24" onClick="move(document.frmCadFeriado.selSala_Selecionado,document.frmCadFeriado.selSala,1)"></div></td>
				</tr>
				<tr>
				  <td height="26" align="center" valign="middle"></td>
				</tr>				
		  </table>
		 
		  <input type="hidden" value="<%=strMostraSala%>" name="txtMostraSala">	
		  
	</form>
	</body>
	<%	
	rsSala.close
	set rsSala = nothing
	
	db_banco.close
	set db_banco = nothing
	%>
	<script language="javascript">		
		//alert(document.frmCadFeriado.rdFeriNacional(1).checked);
			
		//if ((document.frmCadFeriado.txtMostraSala.value == 'Não') || (document.frmCadFeriado.rdFeriNacional(1).checked != true))
		if (document.frmCadFeriado.rdFeriNacional(1).checked != true)
		{
			document.frmCadFeriado.selCorte.disabled = true;	
			//document.frmCadFeriado.selCT.disabled = true;
			document.frmCadFeriado.selSala.disabled = true;
			document.frmCadFeriado.selSala_Selecionado.disabled = true;			
		}
		else
		{
			document.frmCadFeriado.selCorte.disabled = false;	
			//document.frmCadFeriado.selCT.disabled = false;
			document.frmCadFeriado.selSala.disabled = false;
			document.frmCadFeriado.selSala_Selecionado.disabled = false;
		}				
	</script>

</html>
