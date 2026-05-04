<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
'db_banco.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("Petrobras 2004_v2.mdb")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

if Request("selCorte") <> "" then
	Session("Corte") = Request("selCorte")
end if	

if Request("txtNomeSala") <> "" then
	strNomeSala = Request("txtNomeSala")
else
	strNomeSala = ""
end if			

if Request("txtCapacidade") <> "" then
	strCapacidade = Request("txtCapacidade")
else
	strCapacidade = ""
end if		
					
strAcao 		= trim(Request("pAcao"))
strSala	 		= trim(Request("selSala"))
strCdCT 		= ""
	
if strAcao = "I" then
	strNomeAcao = "Inclusão"
elseif strAcao = "A" then 
	strNomeAcao ="Alteração"
end if 

strUC = trim(Request("pUC"))

strSQLCT = ""
strSQLCT = strSQLCT & "SELECT CTRO_CD_CENTRO_TREINAMENTO, CTRO_TX_NOME_CENTRO_TREINAMENTO "
strSQLCT = strSQLCT & "FROM GRADE_CENTRO_TREINAMENTO " 
strSQLCT = strSQLCT & "WHERE CORT_CD_CORTE = " & Session("Corte") 
strSQLCT = strSQLCT & " ORDER BY CTRO_TX_NOME_CENTRO_TREINAMENTO"
'Response.write strSQLCT
'Response.end
set rsCT = db_banco.Execute(strSQLCT)

if strAcao = "A" then	

	strSQLAltSala = ""
	strSQLAltSala = strSQLAltSala & "SELECT CTRO_CD_CENTRO_TREINAMENTO, SALA_TX_NOME_SALA, SALA_NUM_CAPACIDADE, SALA_CD_UC "
	strSQLAltSala = strSQLAltSala & "FROM GRADE_SALA "
	strSQLAltSala = strSQLAltSala & "WHERE SALA_CD_SALA = " & strSala	
	strSQLAltSala = strSQLAltSala & " AND CORT_CD_CORTE = " & Session("Corte") 
	'Response.write strSQLAltSala
	'Response.end
	
	Set rdsAltSala = db_banco.Execute(strSQLAltSala)			
	
	if not rdsAltSala.EOF then
		strCdCT 		= rdsAltSala("CTRO_CD_CENTRO_TREINAMENTO")
		strNomeSala 	= rdsAltSala("SALA_TX_NOME_SALA")		
		strCapacidade	= rdsAltSala("SALA_NUM_CAPACIDADE")
		strUC			= rdsAltSala("SALA_CD_UC")
	end if
	
	rdsAltSala.close
	set rdsAltSala = nothing	
end if
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		
		<script language="javascript">
		
			function Verifica_Dif_Numeros(strValor,strNome)	
			{		
				if (isNaN(strValor))
				{									
					if (strNome == 'txtCapacidade')
					{
						alert("O contéudo do campo Capacidade deve ser preenchido apenas com nº!");
						document.frmCadSala.txtCapacidade.value = '';
						document.frmCadSala.txtCapacidade.focus();
						return;
					}
				}
			}		
			
			function Verifica_Capacidade(strValor)
			{
				if (strValor > 20)
					{
						alert("A capacidade máxima por sala é de 20 lugares!");
						document.frmCadSala.txtCapacidade.value = '';
						document.frmCadSala.txtCapacidade.focus();
						return;
					}
			}	
			
			function Confirma()
			{		
				var strTipoAcao = document.frmCadSala.parAcao.value;
						
				if (strTipoAcao == 'A')
				{
					if(document.frmCadSala.txtSala.value == "")
					{
					alert("É necessário o preenchimento do campo SALA!");
					document.frmCadSala.txtSala.focus();
					return;
					}
				}
			
				if(document.frmCadSala.selCT.selectedIndex == 0)
				{
				alert("Selecione um CENTRO DE TREINAMENTO!");
				document.frmCadSala.selCT.focus();
				return;
				}				
								
				if(document.frmCadSala.txtCapacidade.value == "")
				{
				alert("É necessário o preenchimento do campo CAPACIDADE!");
				document.frmCadSala.txtCapacidade.focus();
				return;
				}			
				
				var strValorUniv = 'N';
				if (document.frmCadSala.checkUniversidade.checked == true)
				{
					strValorUniv = 'S';
				}
				
				document.frmCadSala.action="grava_sala.asp?pUniv="+strValorUniv;
				document.frmCadSala.submit();			
			}
			
			function submet_pagina(strValor, strTipo)
			{			
				var strNomeSala = document.frmCadSala.txtNomeSala.value;
				var strCapacidade = document.frmCadSala.txtCapacidade.value;
											
				var strValorUC = 'N';
				if (document.frmCadSala.checkUniversidade.checked == true)
				{
					strValorUC = 'S';
				}			
												
				window.location.href = "cadastra_sala.asp?selCT=0&selCorte="+strValor+"&txtNomeSala="+strNomeSala+"&txtCapacidade="+strCapacidade+"&pUC="+strValorUC;
			}
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmCadSala">
		
			<input type="hidden" value="<%=strSala%>" name="hdSala"> 
			<input type="hidden" value="<%=strAcao%>" name="parAcao"> 
			<input type="hidden" value="Sala" name="parTipo"> 
				   
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
					<div align="center"><font face="Verdana" color="#330099" size="3"><b> Salas - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="725" height="233">
			  
			    <tr>
			  	  <td height="27"></td>			  	 
			  	  <td height="27" valign="middle" align="left" colspan="2"><font face="Verdana" color="#330099" size="2"><b>Operação:</b>&nbsp;&nbsp;<%=strNomeAcao%></font></td>
		  	    </tr>
								
			  	<tr>
			  	  <td height="9"></td>
			  	  <td height="9" valign="middle" align="left"></td>
			  	  <td height="9" valign="middle" align="left"></td>
		  	    </tr>
			  
			  <tr>			
					 <td width="139" height="36"></td>						
					<td width="231" height="36" valign="middle" align="left">
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
					 <td height="36" valign="middle" align="left" width="341"> 							
														   
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
						rsCorte.close
						set rsCorte = nothing			 
						%>				  </td>
			    </tr>
				
			  	<tr> 
				  <td width="139" height="42"></td>
				  <td width="231" height="42" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Nome da Sala:</b></font></td>
				  <td height="42" valign="middle" align="left" width="341"> 
					<%'if strSala <> "" and strAcao = "A" then%> 						
						<input type="hidden" name="txtSala" value="<%=strSala%>">	
					    <input type="text" name="txtNomeSala" maxlength="100" size="50" value="<%=strNomeSala%>">
					    <%'else%>
						<!--<input type="text" name="txtSala" maxlength="100" size="50" value="<%'=strNomeSala%>">	-->
				  <%'end if%>				  </td>
				</tr>
				  
				<tr> 
				  <td width="139" height="36"></td>
				  <td width="231" height="36" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Centro de Treinamento:</b></font></td>
				  <td height="36" valign="middle" align="left" width="341"> 
				   
					 <select name="selCT" size="1" tabindex="1">
						<option value="0">-- Escolha um centro de treinamento ---</option>
							<%
							 do while not rsCT.eof
							 	if strCdCT = rsCT("CTRO_CD_CENTRO_TREINAMENTO") then
									%>
									<option value="<%=rsCT("CTRO_CD_CENTRO_TREINAMENTO")%>" selected><%=rsCT("CTRO_TX_NOME_CENTRO_TREINAMENTO")%></option>
									<%
								else
									%>
									<option value="<%=rsCT("CTRO_CD_CENTRO_TREINAMENTO")%>"><%=rsCT("CTRO_TX_NOME_CENTRO_TREINAMENTO")%></option>
									<%
								end if
								
								rsCT.MoveNext
							 Loop
							 %>
				    </select>
								  
				  </td>
				</tr>
												
				<tr> 
				  <td width="139" height="32"></td>
				  <td width="231" height="32" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Capacidade:</b></font></td>
				  <td height="32" valign="middle" align="left" width="341"> 
					<input type="text" name="txtCapacidade" maxlength="5" size="5" value="<%=strCapacidade%>" onKeyUp="javascript:Verifica_Dif_Numeros(this.value,this.name);" onChange="javascript:Verifica_Capacidade(this.value);">					
				  </td>
				</tr>   
								
				<tr> 
				  <td width="139" height="32"></td>
				  <td width="231" height="32" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Universidade Corporativa:</b></font></td>
				  <td height="32" valign="middle" align="left" width="341"> 
					<%if strUC = "S" then%>
						<input type="checkbox" name="checkUniversidade" checked>						
					<%else%>
                    	<input type="checkbox" name="checkUniversidade">					
                    <%end if%>				
				  </td>
				</tr>   				
								
				<tr> 
				  <td width="139" height="1"></td>
				  <td width="231" height="1" valign="middle" align="left"></td>
				  <td height="1" valign="middle" align="left" width="341"> </td>
				</tr>   
		  </table>
		</form>
	</body>
</html>
<%
db_banco.close
set db_banco = nothing
%>
