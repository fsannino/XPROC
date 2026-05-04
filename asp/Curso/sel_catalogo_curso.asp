
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("rdbStatus") <> "" then
	strStatus = request("rdbStatus")
else
	strStatus = "0"
end if

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
set rs_mega = db.execute(str_SQL_MegaProc)

set rs_onda = db.execute("SELECT * FROM " & Session("PREFIXO") & "ABRANGENCIA_CURSO WHERE ONDA_CD_ONDA<>4 ORDER BY ONDA_TX_DESC_ONDA")
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>	
		<script>
			function Confirma()
			{		
				if (document.frm1.rdbStatus[0].checked)
				{ 
					document.frm1.hdStatus.value = 0;
				}
				
				if (document.frm1.rdbStatus[1].checked)
				{
					document.frm1.hdStatus.value = 1;
				}
				
				if (document.frm1.rdbStatus[2].checked)
				{
					document.frm1.hdStatus.value = 2;
				}
			
				document.frm1.action = 'rel_catalogo_curso.asp';
				document.frm1.submit();
			}
		</script>
	</head>	
	
	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" action="" name="frm1">
			<input type="hidden" name="txtImp" size="20">
			<input type="hidden" name="txtQua" size="20">
			<input type="hidden" name="hdStatus">
			
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
						<div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
					  </td>
					</tr>
				  </table>
				</td>
			  </tr>
			  <tr bgcolor="#00FF99">
				<td colspan="3" height="20">
				  <table width="625" border="0" align="center">
					<tr>
						<td width="26"><a href="javascript:Confirma()"><img border="0" src="../Funcao/confirma_f02.gif" width="24" height="24"></a></td>
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
			  <td>
			  </td>
			</tr>
			<tr>
			  <td>
				<div align="center"><font face="Verdana" color="#330099" size="3">Seleção Católogo de Cursos</font></div>
			  </td>
			</tr>
			<tr>
			  <td>&nbsp;</td>
			</tr>
		  </table>
		  <table border="0" width="868" height="45">
			  <tr>				
		  		<td width="100" height="29"></td>				
		  		<td width="192" height="29" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo:</b></font></td>
				<td width="543" height="29" valign="middle" align="left"> 
					<select size="1" name="selMegaProcesso">
						<option value="0">== TODOS ==</option>
						<%
						do until rs_mega.eof=true%>						
							<option value="<%=rs_mega("MEPR_CD_MEGA_PROCESSO")%>"><%=rs_mega("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
							<%							
							rs_mega.movenext
						loop
						
						rs_mega.close
						set rs_mega = nothing
						%>
					</select>
				</td>				
			</tr>
			<tr>				
			  <td width="100" height="1"></td>				
			  <td width="192" height="1" valign="middle" align="left"></td>				
			  <td height="1" valign="middle" align="left" width="543"></td>
			  </tr>
			  <tr>
				<td height="1"></td>
				<td height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Abrang&ecirc;ncia do Curso:</b></font></td>
				<td height="1" valign="middle" align="left">
					<select size="1" name="selOnda">
						<option value="0">== TODAS ==</option>
						<%
						do until rs_onda.eof = true%>
							%>
							<option value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
							<%
							rs_onda.movenext
						loop
						
						rs_onda.close
						set rs_onda = nothing
						%>
					</select>
				</td>
			  </tr>
			  <tr>
				<td height="1"></td>
				<td height="1" valign="middle" align="left">&nbsp;</td>
				<td height="1" valign="middle" align="left"> <font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Caso o curso seja pertinente &agrave;s solu&ccedil;&otilde;es antecipada e definitiva, &nbsp;marcar a op&ccedil;&atilde;o &quot;PET&quot; adequada.</font></td>
			  </tr> 
			  <tr>
		  	  <td width="100" height="1"></td>
			  <td height="25" valign="middle"><font face="Verdana" size="2" color="#330099"><b>Status Curso:</b></font></td>
			  <td height="25" valign="middle" colspan="2">
	  
			  <%if strStatus = "0" then%>	
				<input name="rdbStatus" type="radio" value="0" checked>&nbsp;<font face="Verdana" size="2" color="#330099">Todos</font>
			  <%else%>	
				<input name="rdbStatus" type="radio" value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Todos</font>
			  <%end if%>
			  
			  <%if strStatus = "1" then%>	  
				<input name="rdbStatus" type="radio" value="0" checked>&nbsp;<font face="Verdana" size="2" color="#330099">Ativo</font>&nbsp;&nbsp;
			  <%else%>	
				<input name="rdbStatus" type="radio" value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Ativo</font>&nbsp;&nbsp;
			  <%end if%>
			  
			  <%if strStatus = "2" then%>	
				<input name="rdbStatus" type="radio" value="0" checked>&nbsp;<font face="Verdana" size="2" color="#330099">Inativo</font>&nbsp;&nbsp;
			  <%else%>	
				<input name="rdbStatus" type="radio" value="0">&nbsp;<font face="Verdana" size="2" color="#330099">Inativo</font>&nbsp;&nbsp;
			  <%end if%>	 	 
			 
			  </td>      
		  </tr>
			</table>
		</form>
	</body>
</html>
