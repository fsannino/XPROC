<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
strAcao = trim(Request("pAcao"))
'Response.write strAcao & "<br>"

intCdCorte = Request("selCorte")
session("Corte") = intCdCorte

if strAcao = "I" then
	strNomeAcao = "Inclusão"
elseif strAcao = "A" then 
	strNomeAcao ="Alteração"
end if  

strSQLLocalidade = ""
strSQLLocalidade = strSQLLocalidade & "SELECT LOC_CD_LOCALIDADE, LOC_TX_NOME_LOCALIDADE "
strSQLLocalidade = strSQLLocalidade & "FROM GRADE_LOCALIDADE " 
strSQLLocalidade = strSQLLocalidade & "ORDER BY LOC_TX_NOME_LOCALIDADE"
'Response.write strSQLLocalidade
'Response.end
set rsLocalidade = db_banco.Execute(strSQLLocalidade)

if strAcao = "A" then	

	intCdCT = trim(Request("selCT"))
	
	strSQLAltCT = ""
	strSQLAltCT = strSQLAltCT & "SELECT CTRO_CD_CENTRO_TREINAMENTO, LOC_CD_LOCALIDADE, CTRO_TX_NOME_CENTRO_TREINAMENTO "
	strSQLAltCT = strSQLAltCT & "FROM GRADE_CENTRO_TREINAMENTO " 
	strSQLAltCT = strSQLAltCT & "WHERE CTRO_CD_CENTRO_TREINAMENTO = "& intCdCT
	strSQLAltCT = strSQLAltCT & " AND CORT_CD_CORTE = " & intCdCorte 
	'Response.write strSQLAltCT
	'Response.end
	
	Set rdsAltCT = db_banco.Execute(strSQLAltCT)			
	
	if not rdsAltCT.EOF then	
		strNomeCT	= rdsAltCT("CTRO_TX_NOME_CENTRO_TREINAMENTO")		
		intCDLocalidade = rdsAltCT("LOC_CD_LOCALIDADE")	
	end if
	
	rdsAltCT.close
	set rdsAltCT = nothing	
end if
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		
		<script language="javascript">
			function Confirma()
			{				
				if(document.frmCadCT.txtNomeCT.value == "")
				{
					alert("É necessário o preenchimento do campo CENTRO DE TREINAMENTO!");
					document.frmCadCT.txtNomeCT.focus();
					return;
				}				
				
				if(document.frmCadCT.selLocalidade.selectedIndex == 0)
				{
					alert("É necessário a seleção de uma LOCALIDADE!");
					document.frmCadCT.selLocalidade.focus();
					return;
				}				
				
				document.frmCadCT.action="grava_CT.asp";
				document.frmCadCT.submit();			
			}
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmCadCT">		  
			 
			<input type="hidden" value="<%=strAcao%>" name="parAcao"> 				
			
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
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Centro de Treinamento - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="695" height="146">
			  
			  	<tr>
			  	  <td height="20"></td>
			  	  <td height="20" valign="middle" align="center" colspan="2">
				  	<%if strGrava = "GravaTurma" then%>
				  	<font face="Verdana" color="#FE5A31" size="2"><b><%=strMSG%></b></font>
					<%end if%>
				  </td>			  	  
		  	    </tr>
			  	<tr>
			  	  <td height="31"></td>			  	 
			  	  <td height="31" valign="middle" align="left" colspan="2"><font face="Verdana" color="#330099" size="2"><b>Operação:</b>&nbsp;&nbsp;<%=strNomeAcao%></font></td>
		  	    </tr>
			  	<tr>
			  	  <td height="7"></td>
			  	  <td height="7" valign="middle" align="left"></td>
			  	  <td height="7" valign="middle" align="left"></td>
		  	    </tr>
				
				 <tr>
					 <td height="32" colspan="1"></td>
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
					   <select name="selCorte" size="1">							
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
				
				<tr> 
				  <td width="151" height="42"></td>
				  <td width="192" height="42" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Centro de Treinamento:</b></font></td>
				  <td height="42" valign="middle" align="left" width="338">				  	
					 <input type="hidden" name="txtCdCT" value="<%=intCdCT%>">					
	 	      	     <input type="text" name="txtNomeCT" maxlength="50" size="50" value="<%=strNomeCT%>">				  
				  </td>
				</tr> 
								
				<tr> 
				  <td width="151" height="1"></td>
				  <td width="192" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Localidade:</b></font></td>				
				  <td height="1" valign="middle" align="left" width="338"> 
					<select name="selLocalidade" size="1" tabindex="1">
						<option value="0">-- Escolha uma localidade ---</option>
						<%
						 do while not rsLocalidade.eof
						 
						 	  if intCDLocalidade = rsLocalidade("LOC_CD_LOCALIDADE") then
								  %>												  
								  <option value="<%=rsLocalidade("LOC_CD_LOCALIDADE")%>" selected><%=Ucase(rsLocalidade("LOC_TX_NOME_LOCALIDADE"))%></option>
								  <%
							  else
							  	 %>												  
							  	<option value="<%=rsLocalidade("LOC_CD_LOCALIDADE")%>"><%=Ucase(rsLocalidade("LOC_TX_NOME_LOCALIDADE"))%></option>
							 	 <%
							  end if
							  rsLocalidade.MoveNext
						 Loop
						 %>
					</select>
					<%
					rsLocalidade.close
					set rsLocalidade = nothing
					%>					
				  </td>	
				</tr>
				<tr> 
				  <td width="151" height="1"></td>
				  <td width="192" height="1" valign="middle" align="left"></td>
				  <td height="1" valign="middle" align="left" width="338"> </td>
				</tr>   
		  </table>
	</form>
	</body>
	<%	
	db_banco.close
	set db_banco = nothing
	%>
</html>
