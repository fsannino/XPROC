<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")
		
str_Acao = "I"
		
Public Function RetornaNomeUsuario(strChave, strTipoResponsavel)		
	sql_VerUsuario= ""				
	sql_VerUsuario = sql_VerUsuario & " SELECT USUA_TX_NOME_USUARIO"		
	sql_VerUsuario = sql_VerUsuario & " FROM XPEP_EQUIPE_SINERGIA "
	sql_VerUsuario = sql_VerUsuario & " WHERE USUA_TX_CD_USUARIO = '" & strChave & "'"					
					
	set rds_VerUsuario = db_Cogest.Execute(sql_VerUsuario)
	
	if not rds_VerUsuario.eof then									
		RetornaNomeUsuario = Ucase(rds_VerUsuario("USUA_TX_NOME_USUARIO"))					
	else
		sql_VerUsuarioLegado = ""
		sql_VerUsuarioLegado = sql_VerUsuarioLegado & " SELECT USMA_TX_NOME_USUARIO"		
		sql_VerUsuarioLegado = sql_VerUsuarioLegado & " FROM USUARIO_MAPEAMENTO "
		sql_VerUsuarioLegado = sql_VerUsuarioLegado & " WHERE USMA_TX_MATRICULA <> 0"
		sql_VerUsuarioLegado = sql_VerUsuarioLegado & " AND USMA_CD_USUARIO = '" & strChave & "'"					
		set rds_VerUsuarioLegado = db_Cogest.Execute(sql_VerUsuarioLegado)
		
		if not rds_VerUsuarioLegado.eof then
			RetornaNomeUsuario = Ucase(rds_VerUsuarioLegado("USMA_TX_NOME_USUARIO"))
		else
			RetornaNomeUsuario = "USUÁRIO NÃO LOCALIZADO."
		end if
		rds_VerUsuarioLegado.close
		set rds_VerUsuarioLegado = nothing
	end if		
	rds_VerUsuario.close
	set rds_VerUsuario = nothing
End function

strChaveUsuario 	= Request("pChaveUsua")	
strCampo 			= Request("pCampo")	
		
if strCampo = "txtUsuarioAcesso" then
	strUsuaAcesso = " - " & RetornaNomeUsuario(strChaveUsuario,"")
	strUsuarioAcesso = RetornaNomeUsuario(strChaveUsuario,"")
end if				

if strChaveUsuario <> "" then
	sqlConsUsuario = ""
	sqlConsUsuario = sqlConsUsuario & "SELECT USUA_CD_USUARIO, USUA_TX_CATEGORIA "
	sqlConsUsuario = sqlConsUsuario & "FROM GRADE_USUARIO "
	sqlConsUsuario = sqlConsUsuario & "WHERE USUA_CD_USUARIO ='" & strChaveUsuario & "'"				
	set rst_ConsUsuario = db_Cogest.Execute(sqlConsUsuario)
		
	if not rst_ConsUsuario.eof then
		str_Acao = "A"
		
		strCategoria = rst_ConsUsuario("USUA_TX_CATEGORIA")											
		select case strCategoria
			case "A":
				strChecked_A = "checked"
			case "B":
				strChecked_B = "checked"
			case "C":
				strChecked_C = "checked"
			case "D":
				strChecked_D = "checked"
		end select				
		
	else				
		strCategoria = ""		
	end if	
	
	rst_ConsUsuario.close
	set rst_ConsUsuario = nothing
else
	 strCategoria = ""
end if

if str_Acao = "I" then
	str_Texto_Acao = "Inclusão"
elseif str_Acao = "A" then
	str_Texto_Acao = "Alteração"
end if	
%>	

<html>
	<head>
	
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

	<script>			
		function confirma_usuario() 
		{ 		
			if (document.frm_Usuario.txtUsuarioAcesso.value == "")
			 { 
			 alert("O campo Chave do Usuário deve ser preenchido!");
			 document.frm_Usuario.txtUsuarioAcesso.focus();
			 return;
			 }	
			 else
			 {
				if (document.frm_Usuario.pNomeUsua.value == "USUÁRIO NÃO LOCALIZADO.")
				{
					 alert("Informe a chave de um Usuário Existente!");
					 document.frm_Usuario.txtUsuarioAcesso.focus();
					 return;
				}
			 }					
			  
			if ((!document.frm_Usuario.rdbCategoria[0].checked)&&
				(!document.frm_Usuario.rdbCategoria[1].checked)&&
				(!document.frm_Usuario.rdbCategoria[2].checked)&&
				(!document.frm_Usuario.rdbCategoria[3].checked))
			{
				 alert("A seleção da Categoria é obrigatória !");
				return;
			 }					
			 
			document.frm_Usuario.action='grava_usuario.asp';
			document.frm_Usuario.submit();			 
		}
					
		function Localiza_Usuario(strCampo)
		{
			if (strCampo == 'txtUsuarioAcesso')
			{
				strUsuario = document.frm_Usuario.txtUsuarioAcesso.value;		
			
				if (strUsuario == '')
				{			
					alert("É obrigatório o preenchimento do campo Chave do Usuário!");
					document.frm_Usuario.txtUsuarioAcesso.focus();
					return;
				}
			}	
			
			document.frm_Usuario.pChaveUsua.value = strUsuario.toUpperCase();			
			document.frm_Usuario.action='cadastra_usuario.asp?pCampo=' + strCampo;
			document.frm_Usuario.submit();			
		}		
				
		function confirma_exclusao()
		{			
		  if(confirm("Confirma a exclusão deste Registro?"))
		  {
			document.frm_Usuario.pAcao.value = 'E';			
			document.frm_Usuario.action='grava_usuario.asp'; 			        
			document.frm_Usuario.submit();
		  }
		}	
	</script>
</head>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">

<form name="frm_Usuario" method="POST">
	
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
					<td width="26"><a href="javascript:confirma_usuario()"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
				  <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
				  <td width="26">&nbsp;</td>
				  <td width="195"></td>
					 <td width="28"></td>  
						<td width="250"></td>
				  <td width="28"></td>
				  <td width="26">&nbsp;</td>
				  <td width="159"></td>
				</tr>
			  </table>
			</td>
		  </tr>
  </table>
	<%if strUsuaAcesso <> "" then%>	
		<input type="hidden" value="<%=strUsuaAcesso%>" name="hdUsuaAcesso">
    <%else%>	
  		<input type="hidden" value="<%=Request("hdUsuaAcesso")%>" name="hdUsuaAcesso">
	<%end if%>	 
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">		
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
	    </tr>
		<tr> 
		  <td width="36%">&nbsp;</td>
		  <td width="50%"><font face="Verdana" color="#330099" size="3"><b>Cadastro de Usu&aacute;rio</b></font></td>
		  <td width="14%">&nbsp;</td>
		</tr>
	  </table>
	  <table width="101%" border="0" cellspacing="0" cellpadding="0" height="172">
		<tr> 
		  <td width="23%" height="21">&nbsp;</td>
		  <td width="17%" height="21">&nbsp;</td>
		  <td width="59%" height="21">&nbsp;</td>
	    </tr>
		
		<tr>
		  <td width="23%" height="27">&nbsp;</td>		    	 
		  <td height="27" valign="middle" align="left" colspan="2"><font face="Verdana" color="#330099" size="2"><b>Operação:</b>&nbsp;&nbsp;<%=str_Texto_Acao%></font></td>
		</tr>
								
		<tr> 
		  <td width="23%" height="15"></td>
		  <td width="17%" height="15"></td>
		  <td width="59%" height="15"></td>
	    </tr>
								
		<tr> 
		  <td width="23%" height="52"></td>
		  <td width="17%" height="52"><font face="Verdana" color="#330099" size="2"><b>Chave do Usu&aacute;rio:</b></font></td>
		  <td width="59%" height="52"> 	
			<%if Request("txtUsuarioAcesso") <> "" then%>
				<input type="text" name="txtUsuarioAcesso" size="5" maxlength="4" value="<%=Request("txtUsuarioAcesso")%>" onblur="javascript:Localiza_Usuario('txtUsuarioAcesso');">
			<%else%>
				<input type="text" name="txtUsuarioAcesso" size="5" maxlength="4" value="<%=strUsuarioAcesso%>" onblur="javascript:Localiza_Usuario('txtUsuarioAcesso');">
			<%end if%>		
			
			<font face="Verdana" color="#330099" size="2"><b>				
			<%
			if strUsuaAcesso <> "" then
				Response.write strUsuaAcesso
			else
				Response.write Request("hdUsuaAcesso") 
			end if
			%>
		  </b></font>		  </td>
	    </tr>			
		<tr> 
		  <td width="23%" height="45"></td>
		  <td width="17%" height="45" valign="top"><font face="Verdana" color="#330099" size="2"><b>Categoria:</b></font></td>
		  <td width="59%" height="45">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr> 
				<td width="97%"> 
				  <input type="radio" name="rdbCategoria" value="A" <%=strChecked_A%>>
				  <font face="Verdana" color="#330099" size="2">Administrador - Cadastro, Corte, Usu&aacute;rio, Feriado e Consulta.</font></td>
				<td width="3%"></td>
			  </tr>
			  <tr>
				<td>				  
				  <input type="radio" name="rdbCategoria" value="B" <%=strChecked_B%>>
				  <font face="Verdana" color="#330099" size="2">Funcional 1 - Cadastro e Consulta.</font></td>
				<td>&nbsp;</td>
			  </tr>
			  <tr>
				<td>
				  <input type="radio" name="rdbCategoria" value="C" <%=strChecked_C%>>
				  <font face="Verdana" color="#330099" size="2">Funcional 2 - Cadastro, feriado e Consulta.</font></td>
				<td>&nbsp;</td>
			  </tr>
			  <tr> 
				<td width="97%">
					<input type="radio" name="rdbCategoria" value="D" <%=strChecked_D%>>				  
					<font face="Verdana" color="#330099" size="2">Consulta</font></td>
				<td width="3%"></td>
			  </tr>
			</table>
		  </td>
		</tr>
  </table>		
	
	<table width="625" border="0" align="center">
		<tr>
		  <td>&nbsp;</td>
		  <td></td>
		  <td></td>
		  <td></td>
		  <td></td>
		  <td>&nbsp;</td>
		  <td></td>
		</tr>
		<tr>
		  <td width="73" height="37"></td>
		  <td width="30" height="37"></td>
		  <td width="58" height="37">&nbsp;</td>
		  <td width="115" height="37">
			<%if str_Acao = "A" then%>
				<input type="button" value="  Excluir  " onClick="javascript:confirma_exclusao();" class="boton_box">
			<%end if%>	  </td>	  	
		  <td width="80" height="37"></td>
		  <td width="99" height="37">&nbsp;</td>
		  <td width="140" height="37"><div align="center"></div></td>
		</tr>
  </table>
	  <input type="hidden" value="<%=str_Acao%>" name="pAcao">		  		 
	  <input type="hidden" value="" name="pChaveUsua">
	  <input type="hidden" value="<%=strUsuarioAcesso%>" name="pNomeUsua">
	  <input type="hidden" name="pOndaSelecionada">		  
</form>
	<%
		db_Cogest.close
		set db_Cogest = nothing
	%>
</body>
</html>