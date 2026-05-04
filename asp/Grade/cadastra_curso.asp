<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
'db_banco.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("Petrobras 2004_v2.mdb")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
strCdCurso	 = trim(Request("selCurso"))
		
if trim(Request("selCorte")) <> "" then
	Session("Corte") = trim(Request("selCorte"))
end if

strSQLCurso = ""
strSQLCurso = strSQLCurso & "SELECT CURS_TX_NOME_CURSO, CURS_TX_CENTRALIZADO, CURS_TX_IN_LOCO "
strSQLCurso = strSQLCurso & "FROM GRADE_CURSO "
strSQLCurso = strSQLCurso & "WHERE CURS_CD_CURSO = '" & strCdCurso & "' "
strSQLCurso = strSQLCurso & "AND CORT_CD_CORTE = " & Session("Corte")
'Response.write strSQLCurso
'Response.end
Set rdsCurso = db_banco.Execute(strSQLCurso)		

strNomeCurso = rdsCurso("CURS_TX_NOME_CURSO")	
strNomeDesc = trim(rdsCurso("CURS_TX_CENTRALIZADO"))
strNomeInloco = trim(rdsCurso("CURS_TX_IN_LOCO"))

'************ UNIDADES ****************
strSQLUnidade = ""
strSQLUnidade = strSQLUnidade & "SELECT UNID_CD_UNIDADE, UNID_TX_DESC_UNIDADE "
strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE " 
strSQLUnidade = strSQLUnidade & "WHERE CORT_CD_CORTE = " & Session("Corte")

'*** PARA RETIRAR OS CADASTRADOS DA LISTA DE SELEÇÃO
strSQLUnidade = strSQLUnidade & " AND UNID_CD_UNIDADE NOT IN "
strSQLUnidade = strSQLUnidade & "(SELECT UNIDADE.UNID_CD_UNIDADE "
strSQLUnidade = strSQLUnidade & "FROM GRADE_UNIDADE UNIDADE, GRADE_CURSO_UNIDADE CURSO_UNID "
strSQLUnidade = strSQLUnidade & "WHERE UNIDADE.UNID_CD_UNIDADE = CURSO_UNID.UNID_CD_UNIDADE "
strSQLUnidade = strSQLUnidade & " AND CURSO_UNID.CURS_CD_CURSO = '" & strCdCurso & "'"
strSQLUnidade = strSQLUnidade & " AND UNIDADE.CORT_CD_CORTE = " & Session("Corte") 
strSQLUnidade = strSQLUnidade & " AND CURSO_UNID.CORT_CD_CORTE = " & Session("Corte") & ") "	

strSQLUnidade = strSQLUnidade & " ORDER BY UNID_TX_DESC_UNIDADE"
'Response.write strSQLUnidade & "<br><br><br>"
'Response.end
set rsUnidade = db_banco.Execute(strSQLUnidade)

'************ UNIDADES x CURSO ****************
strSQLUnidadeCurso = ""
strSQLUnidadeCurso = strSQLUnidadeCurso & "SELECT UNIDADE.UNID_CD_UNIDADE, UNIDADE.UNID_TX_DESC_UNIDADE "
strSQLUnidadeCurso = strSQLUnidadeCurso & "FROM GRADE_UNIDADE UNIDADE, GRADE_CURSO_UNIDADE CURSO_UNID "
strSQLUnidadeCurso = strSQLUnidadeCurso & "WHERE UNIDADE.UNID_CD_UNIDADE = CURSO_UNID.UNID_CD_UNIDADE "
strSQLUnidadeCurso = strSQLUnidadeCurso & " AND CURSO_UNID.CURS_CD_CURSO = '" & strCdCurso & "'"
strSQLUnidadeCurso = strSQLUnidadeCurso & " AND UNIDADE.CORT_CD_CORTE = " & Session("Corte") 
strSQLUnidadeCurso = strSQLUnidadeCurso & " AND CURSO_UNID.CORT_CD_CORTE = " & Session("Corte") 
'Response.write strSQLUnidadeCurso
'Response.end

Set rdsUnidadeCurso = db_banco.Execute(strSQLUnidadeCurso)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		<script language="javascript" src="../js/troca_lista.js"></script>
		
		<script language="javascript">	
			
			function Confirma()
			{			
				if ((document.frmCadCurso.rdDescentralizado(0).checked == false)&&(document.frmCadCurso.rdDescentralizado(1).checked == false))
				{
					alert('Indforme se o curso é Centralizado ou Descentralizado');
					document.frmCadCurso.rdDescentralizado.focus();
					return;
				}
				
				if ((document.frmCadCurso.rdInLoco(0).checked == false)&&(document.frmCadCurso.rdInLoco(1).checked == false))
				{
					alert('Indforme se o curso é In Loco ou não');
					document.frmCadCurso.rdDescentralizado.focus();
					return;
				}
				
				//*** Monta uma string com os CURSOS Selecionados, separados por vírgula
				carrega_txt(document.frmCadCurso.selUnidade_Selecionado)										
																		
				document.frmCadCurso.action="grava_curso.asp";
				document.frmCadCurso.submit();			
			}
			
			function carrega_txt(fbox) 
			{
				document.frmCadCurso.txtUnid_Selecionadas.value = '';
				for(var i=0; i<fbox.options.length; i++) 
				{
					if (i == 0)
					{
						document.frmCadCurso.txtUnid_Selecionadas.value = fbox.options[i].value;
					}
					else
					{					
						document.frmCadCurso.txtUnid_Selecionadas.value = document.frmCadCurso.txtUnid_Selecionadas.value + "," + fbox.options[i].value;
					}
				}
			}	
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmCadCurso">
		
			<input type="hidden" name="txtUnid_Selecionadas">	
			<input type="hidden" value="<%=strCdCurso%>" name="hdCurso"> 
							   
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
						<td width="26"><a href="javascript:Confirma();"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
					  <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
					  <td width="26">&nbsp;</td>
					  <td width="195"></td>
						<td width="27"></td>  <td width="50"></td>
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
				  <td height="10">
				  </td>
				</tr>
				<tr>
				  <td>
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Curso - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="725" height="180">
													
				<tr>
					 <td height="21" colspan="1"></td>
					 <td width="207" valign="middle">						
					   <font face="Verdana" size="2" color="#330099"><b>Corte:&nbsp;</b></font>
					 </td>
					 <td width="300" colspan="2" valign="middle">					   
					 	<%					 
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
						rsCorte.close
						set rsCorte = nothing		
						%>				   	   
					</td>
			    </tr>			
							
			  	<tr> 
				  <td width="171" height="26"></td>
				  <td width="188" height="26" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Código do Curso:</b></font></td>
				  <td height="26" valign="middle" align="left" width="352"><font face="Verdana" size="2" color="#330099"><%=strCdCurso%></font>				  </td>
				</tr> 
				
				<tr> 
				  <td width="171" height="22"></td>
				  <td width="188" height="22" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Nome do Curso:</b></font></td>
				  <td height="22" valign="middle" align="left" width="352"><font color="#330099" size="2" face="Verdana"><%=strNomeCurso%></font>				  </td>
				</tr> 					
				
				<tr> 
				  <td width="195" height="22"></td>
				  <td width="172" height="22" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Tipo do Curso:</b></font></td>
				  <td height="22" valign="middle" align="left" width="468">	
				  
				  <%				  
				  if strNomeDesc = "CENTRALIZADO" then
				  %>				  		  	
					<input type="radio" name="rdDescentralizado" checked value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Centralizados</font>&nbsp;&nbsp;
					<input type="radio" name="rdDescentralizado" value="1">&nbsp;<font face="Verdana" size="2" color="#330099">Descentralizados</font>&nbsp;&nbsp;
				  <%				  
				  elseif strNomeDesc = "DESCENTRALIZADO" then
				  %>
					<input type="radio" name="rdDescentralizado" value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Centralizados</font>&nbsp;&nbsp;
					<input type="radio" name="rdDescentralizado" checked value="1">&nbsp;<font face="Verdana" size="2" color="#330099">Descentralizados</font>&nbsp;&nbsp;
				  <%
				  else
				  %>
					<input type="radio" name="rdDescentralizado" value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Centralizados</font>&nbsp;&nbsp;
					<input type="radio" name="rdDescentralizado" value="1">&nbsp;<font face="Verdana" size="2" color="#330099">Descentralizados</font>&nbsp;&nbsp;
				  <%			  
				  end if
				  %>		
				  </td>
				</tr>   
				
				<tr> 
				  <td width="195" height="40"></td>
				  <td width="172" height="40" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>In Loco:</b></font></td>
				  <td height="40" valign="middle" align="left" width="468">			  	
				  <%				 
				  if strNomeInloco = "S" then
				  %>				  		  	
					<input type="radio" name="rdInLoco" checked value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Sim</font>&nbsp;&nbsp;
					<input type="radio" name="rdInLoco" value="1">&nbsp;<font face="Verdana" size="2" color="#330099">Não</font>&nbsp;&nbsp;
				  <%				  
				  elseif strNomeInloco = "N" then
				  %>
					<input type="radio" name="rdInLoco" value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Sim</font>&nbsp;&nbsp;
					<input type="radio" name="rdInLoco" checked value="1">&nbsp;<font face="Verdana" size="2" color="#330099">Não</font>&nbsp;&nbsp;
				  <%
				  else				   
				  %>
					<input type="radio" name="rdInLoco" value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Sim</font>&nbsp;&nbsp;
					<input type="radio" name="rdInLoco" value="1">&nbsp;<font face="Verdana" size="2" color="#330099">Não</font>&nbsp;&nbsp;
				  <%							  
				  end if
				  %>				  </td>
				</tr> 				
		  </table>
		  
		  <table width="866" height="165" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="159" height="55" rowspan="5"></td>
					<td height="20" colspan="3"><font face="Verdana" size="2" color="#330099"><b>Caso o curso seja Roll-Out, selecione uma ou mais unidades:</b></font></td>
				</tr>
				<tr> 
				  <td width="285" height="55" rowspan="5" align="center" valign="middle">        
				  <p align="left">
				  	<font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Unidades Dispon&iacute;veis:</font></p>
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
				  <td width="46" height="32" align="center" valign="middle"><div align="left"></div></td>
				  <td width="376" rowspan="5" align="center" valign="middle">
				  <p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Unidades Selecionadas:</font></p>
					<p align="left">					
						<select name="selUnidade_Selecionado" size="5" multiple>
						<%
						if not rdsUnidadeCurso.EOF then		
							do while not rdsUnidadeCurso.eof					
								%>
								<option value="<%=rdsUnidadeCurso("UNID_CD_UNIDADE")%>"><%=rdsUnidadeCurso("UNID_TX_DESC_UNIDADE")%></option>			
								<%								
								rdsUnidadeCurso.movenext
							loop								
				
							rdsUnidadeCurso.close
							set rdsUnidadeCurso = nothing			
						end if		
						%>					 
						</select>
				  </p></td>
				</tr>
				<tr>
				  <td height="53" align="center" valign="middle"><div align="center"><img src="../../imagens/continua_F01.gif" width="24" height="24" onClick="move(document.frmCadCurso.selUnidade,document.frmCadCurso.selUnidade_Selecionado,1)"></div></td>
				</tr>
				<tr>
				  <td height="34" align="center" valign="middle"><div align="center"><img src="../../imagens/continua2_F01.gif" width="24" height="24" onClick="move(document.frmCadCurso.selUnidade_Selecionado,document.frmCadCurso.selUnidade,1)"></div></td>
				</tr>
				<tr>
				  <td height="26" align="center" valign="middle"></td>
				</tr>				
		  </table>
	</form>
	</body>
</html>
<%
db_banco.close
set db_banco = nothing
%>
