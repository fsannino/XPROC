<%'<!--#include file="../../asp/protege/protege.asp" -->%>
<%
set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3
	
strAcao = trim(Request("pAcao"))
'Response.write strAcao & "<br>"

if strAcao = "I" then
	strNomeAcao = "Inclusão"
elseif strAcao = "A" then 
	strNomeAcao ="Alteração"
end if  

if strAcao = "A" then	

	intCdLocalidade = trim(Request("selLocalidade"))
	
	strSQLAltLocalidade = ""
	strSQLAltLocalidade = strSQLAltLocalidade & "SELECT LOC_CD_LOCALIDADE, LOC_TX_NOME_LOCALIDADE "
	strSQLAltLocalidade = strSQLAltLocalidade & "FROM GRADE_LOCALIDADE " 
	strSQLAltLocalidade = strSQLAltLocalidade & "WHERE LOC_CD_LOCALIDADE = "& intCdLocalidade
	'Response.write strSQLAltLocalidade
	'Response.end
	
	Set rdsAltLocalidade = db_banco.Execute(strSQLAltLocalidade)			
	
	if not rdsAltLocalidade.EOF then	
		strNomeLocalidade	= rdsAltLocalidade("LOC_TX_NOME_LOCALIDADE")			
	end if
	
	rdsAltLocalidade.close
	set rdsAltLocalidade = nothing	
end if
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>		
		<script language="javascript" src="js/digite-cal.js"></script>	
		
		<script language="javascript">
			function Confirma()
			{				
				if(document.frmCadLocalidade.txtNomeLocalidade.value == "")
				{
					alert("É necessário o preenchimento do campo LOCALIDADE!");
					document.frmCadLocalidade.txtNomeLocalidade.focus();
					return;
				}				
				
				document.frmCadLocalidade.action="grava_localidade.asp";
				document.frmCadLocalidade.submit();			
			}
		</script>		
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" name="frmCadLocalidade">		  
			 
			<input type="hidden" value="<%=strAcao%>" name="parAcao"> 	
			<input type="hidden" value="<%=intCdLocalidade%>" name="txtCdLocalidade">		
			
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
					<div align="center"><font face="Verdana" color="#330099" size="3"><b>Localidade - Grade de Treinamento</b></font></div>
				  </td>
				</tr>
				<tr>
				  <td>&nbsp;</td>
				</tr>
			  </table>
			  <table border="0" width="695" height="119">
			  
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
				  <td width="176" height="28"></td>
				  <td width="136" height="28" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Localidade:</b></font></td>
				  <td height="28" valign="middle" align="left" width="369">				  	
					 <input type="hidden" name="txtCdFeriado" value="<%=strCdFeriado%>">					
			 	  <input type="text" name="txtNomeLocalidade" maxlength="50" size="50" value="<%=strNomeLocalidade%>">				  </td>
				</tr> 
								
				<tr> 
				  <td width="176" height="1"></td>
				  <td width="136" height="1" valign="middle" align="left"></td>
				  <td height="1" valign="middle" align="left" width="369"> </td>
				</tr>   
		  </table>
	</form>
	</body>
	<%	
	db_banco.close
	set db_banco = nothing
	%>
</html>
