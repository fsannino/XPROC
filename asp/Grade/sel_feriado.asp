<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
if Request("selCorte") <> "" then
	Session("Corte") =  Request("selCorte")	
end if
	
strSQLFeriado = ""
strSQLFeriado = strSQLFeriado & "SELECT FERI_CD_FERIADO, FERI_TX_NOME_FERIADO, FERI_DT_DATA_FERIADO "
strSQLFeriado = strSQLFeriado & "FROM GRADE_FERIADO " 
strSQLFeriado = strSQLFeriado & "ORDER BY FERI_TX_NOME_FERIADO"
'Response.write strSQLFeriado
'Response.end
set rsFeriado = db_banco.Execute(strSQLFeriado)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		
		<style type="text/css">
		<!--
			.boton_box
			{
				BORDER-RIGHT: black 1px solid;
				BORDER-TOP: black 1px solid;
				BORDER-COLOR: #000066;
				FONT-WEIGHT: bold;
				FONT-SIZE: 12px;
				WORD-SPACING: 2px;
				TEXT-TRANSFORM: capitalize;
				BORDER-LEFT: black 1px solid;
				COLOR: #000066;
				BORDER-BOTTOM: black 1px solid;
				FONT-STYLE: normal;
				FONT-FAMILY: Verdana, Arial, Helvetica, sans-serif;
				BACKGROUND-COLOR: #F1F1F1;
			}
		-->
		</style>

		<script language="javascript" src="js/digite-cal.js"></script>	
		
		<script language="javascript">
			function Confirma(strAcao)
			{										
				if (strAcao =='Incluir')				
				{		
					document.frmFeriado.action = "cadastra_feriado.asp?pAcao=I";
				}
				
				if (strAcao =='Alterar')				
				{		
					if (document.frmFeriado.selFeriado.selectedIndex == 0)
					{
						alert('Selecione um feriado para alteração!');
						document.frmFeriado.selFeriado.focus();
						return;
					}
					else
					{
						document.frmFeriado.action = "cadastra_feriado.asp?pAcao=A";
					}
				}
				
				if (strAcao =='Excluir')				
				{		
					if (document.frmFeriado.selFeriado.selectedIndex == 0)
					{
						alert('Selecione um feriado para exclusão!');
						document.frmFeriado.selFeriado.focus();
						return;
					}
					else
					{											 
					   if (confirm('Por favor confirme a remoção deste feriado.'))
					   {
						 document.frmFeriado.action = "grava_feriado.asp?parAcao=E";
					   }		
					   else
					   {
					   	return;
					   }				
					}
				}
							   
				document.frmFeriado.submit();			
			}
									
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmFeriado">					
			
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
						<td width="24"><!--<a href="javascript:Confirma();"><img border="0" src="../../imagens/confirma_f02.gif"></a>--></td>
					  <td width="46"><!--<font color="#330099" face="Verdana" size="2"><b>Enviar</b></font>--></td>
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
				  <td height="20">
				  </td>
				</tr>
				<tr>
				  <td>
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Cadastro de Feriado </b></font></div>
				  </td>
				</tr>
				
				<tr>
				  <td height="30">
				  </td>
				</tr>			
				
				<tr>
					<td align="center">
						<%			
						if not rsFeriado.eof then
							%>					
							<table border="0" cellpadding="3" cellspacing="3">
								<tr>
								  <td height="30">
								  	<font face="Verdana" color="#330099" size="1">
										Estes são os Feriados cadastrados.<br>
						  		  		Selecione a opção desejada:								  	</font>								  </td>
								</tr>
								<tr>
									<td valign="top" rowspan="3">
										<select name="selFeriado" size="1" tabindex="1">
											<option value="0">-- Escolha um Feriado ---</option>
											<%
											 do while not rsFeriado.eof
												  %>
												  <option value="<%=rsFeriado("FERI_CD_FERIADO")%>"><%=rsFeriado("FERI_TX_NOME_FERIADO") & " - " & Trim(rsFeriado("FERI_DT_DATA_FERIADO"))%></option>
												  <%
												  rsFeriado.MoveNext
											 Loop
											 %>
										</select>
									</td>
										
									<td rowspan="3" width="20">&nbsp;</td>
									<td valign="top" align="left">										
										<input type="submit" Onclick="Confirma('Alterar');" value="Alterar" tabindex="2" class="boton_box">
									</td>
									<td valign="top" align="right">										
										<input type="button" value="Remover" Onclick="Confirma('Excluir');" tabindex="3" class="boton_box">
									</td>
								</tr>
						   
								<tr>
									<td valign="top" colspan="2">										
											<input type="button" value="Incluir novo feriado"  Onclick="Confirma('Incluir');" tabindex="4" class="boton_box">
									</td>
								</tr>
							</table>    
						<%
						else
						%>
							<p><font face="Verdana" color="#330099" size="2"><b>Ainda não existem FERIADOS cadastrados no sistema.</b></p>			   
							<a href="#" Onclick="Confirma('Incluir');">
								<input type="button" value="Incluir novo feriado" tabindex="4" class="boton_box">
							</a>		   
						<%
						end if
						
						rsFeriado.Close
						set rsFeriado = nothing
						%>
					<td>
				</tr>				
			  </table>			
		</form>
	</body>	
	<%	
	db_banco.close
	set db_banco = nothing
	%>	
</html>
