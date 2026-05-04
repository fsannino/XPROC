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
	sqlConsUsuario = sqlConsUsuario & "FROM XPEP_USUARIO "
	sqlConsUsuario = sqlConsUsuario & "WHERE USUA_CD_USUARIO ='" & strChaveUsuario & "'"				
	set rst_ConsUsuario = db_Cogest.Execute(sqlConsUsuario)
	
	sqlConsAcesso = ""
	sqlConsAcesso = sqlConsAcesso & "SELECT ONDA_CD_ONDA "
	sqlConsAcesso = sqlConsAcesso & "FROM XPEP_ACESSO "
	sqlConsAcesso = sqlConsAcesso & "WHERE USUA_CD_USUARIO ='" & strChaveUsuario	& "'"	
	set rst_ConsAcesso = db_Cogest.Execute(sqlConsAcesso)
	
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
		end select					
		
		if not rst_ConsAcesso.eof then
			strAcessoExistentes = ""
			rst_ConsAcesso.movefirst
			do while not rst_ConsAcesso.eof 
				if strAcessoExistentes = "" then
					strAcessoExistentes = rst_ConsAcesso("ONDA_CD_ONDA")
				else
					strAcessoExistentes =  strAcessoExistentes & "," & rst_ConsAcesso("ONDA_CD_ONDA") 
				end if
				rst_ConsAcesso.movenext
			loop
		else
			strAcessoExistentes = ""
		end if
		rst_ConsAcesso.close
		set rst_ConsAcesso = nothing
	else			
		strAcessoExistentes = ""
		'strAcessoExistentes = request("pOndaSelecionada")	
		'if request("rdbCategoria") <> "" then
		'	strCategoria = request("rdbCategoria")
		'	select case strCategoria
		'		case "A":
		'			strChecked_A = "checked"
		'		case "B":
		'			strChecked_B = "checked"
		'		case "C":
		'			strChecked_C = "checked"
		'	end select
		'else
			strCategoria = ""
		'end if
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>:: Cutover - Cadastro de Usuário</title>
<!-- InstanceEndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333; text-decoration: none}
a:hover {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333;  text-decoration: underline}
-->
</style>
<link href="/css/biblioteca.css" rel="stylesheet" type="text/css">
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="Head01" -->

<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaScri" -->
	<script src="../js/troca_lista_sem_retirar.js" language="javascript"></script>		
	<script src="../js/global.js" language="javascript"></script>	
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
				(!document.frm_Usuario.rdbCategoria[2].checked))
			{
				 alert("A seleção da Categoria é obrigatória !");
				return;
			 }		 	
			
			if (document.frm_Usuario.lstOndaSel.options.length == 0)
			 { 
				 alert("É necessária a seleção de pelo menos uma Onda!");
				 document.frm_Usuario.lstOndaSel.focus();
				 return;
			 }				
			 else
			 {
			  carrega_txt(document.frm_Usuario.lstOndaSel);				  
			 }		
			 
			document.frm_Usuario.action='grava_usuario.asp';
			document.frm_Usuario.submit();
			 
		}
					
		function Localiza_Usuario(strCampo)
		{
			carrega_txt(document.frm_Usuario.lstOndaSel);	
		
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
			document.frm_Usuario.action='cad_usuario.asp?pCampo=' + strCampo;
			document.frm_Usuario.submit();			
		}		
		
		function carrega_txt(fbox) 
		{
			document.frm_Usuario.pOndaSelecionada.value = "";
			for(var i=0; i<fbox.options.length; i++) 
			{
				document.frm_Usuario.pOndaSelecionada.value = document.frm_Usuario.pOndaSelecionada.value + "," + fbox.options[i].value;
			}
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
<!-- InstanceEndEditable -->
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<div id="Layer1" style="position:absolute; left:20px; top:10px; width:134px; height:53px; z-index:1"><img src="../img/000005.gif" alt=":: Logo Sinergia" width="134" height="53" border="0" usemap="#Map2"> 
	  <map name="Map2">
	    <area shape="rect" coords="6,7,129,49">
	  </map>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><table width="780" height="44" border="0" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="583" height="44"><img src="../img/_0.gif" width="1" height="1"></td>
	          <td width="197" height="44"><img src="../../../imagens/000043.gif" width="95" height="44"></td>
	        </tr>
	      </table></td>
	  </tr>
</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td bgcolor="#6699CC">
			<table width="780" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td width="154" height="21"><img src="../img/000002.gif" width="154" height="21"></td>
			    <td width="19" height="21"><img src="../img/000003.gif" width="19" height="21"></td>
			    <td width="202" height="21">
					<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
						<strong>
						</strong>
					</font>
			    </td>
			    <td>&nbsp;</td>
		      </tr>
			</table>
	    </td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="5"><img src="../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="780" height="58" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20" height="39"><img src="../img/_0.gif" width="1" height="1"></td>
        <td width="740" height="39" background="../img/000006.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="11%">&nbsp;</td>
              <td width="13%">&nbsp;</td>
              <td width="61%"><font color="#666666" size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>PLANO DE ENTRADA EM PRODU&Ccedil;&Atilde;O</b></font></td>
              <td width="15%"><a href="../../../indexA_xpep.asp"><img src="../img/botao_home_off_01.gif" alt="Ir para tela inicial" width="34" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td width="20" height="39"><img src="../img/_0.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<!-- InstanceBeginEditable name="corpo" -->  
	<form name="frm_Usuario" method="POST">
	
	<%if strUsuaAcesso <> "" then%>	
		<input type="hidden" value="<%=strUsuaAcesso%>" name="hdUsuaAcesso">
	<%else%>
		<input type="hidden" value="<%=Request("hdUsuaAcesso")%>" name="hdUsuaAcesso">
	<%end if%>	 
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">		
		<tr> 
		  <td width="24%">&nbsp;</td>
		  <td width="62%" class="subtitulob">Cadastro de Usu&aacute;rio</td>
		  <td width="14%">&nbsp;</td>
		</tr>
	  </table>
	  <table width="101%" border="0" cellspacing="0" cellpadding="0" height="148">
		<tr> 
		  <td width="2%" height="21">&nbsp;</td>
		  <td width="24%" height="21">&nbsp;</td>
		  <td width="39%" height="21">&nbsp;</td>
		  <td width="35%" height="21">&nbsp;</td>
		</tr>				
		<tr> 
		  <td width="2%" height="25"></td>
		  <td width="24%" height="25" class="campob">Chave do Usu&aacute;rio:</td>
		  <td width="39%" height="25" class="campob"> 	
			<%if Request("txtUsuarioAcesso") <> "" then%>
				<input type="text" name="txtUsuarioAcesso" size="5" maxlength="4" value="<%=Request("txtUsuarioAcesso")%>" onblur="javascript:Localiza_Usuario('txtUsuarioAcesso');">
			<%else%>
				<input type="text" name="txtUsuarioAcesso" size="5" maxlength="4" value="<%=strUsuarioAcesso%>" onblur="javascript:Localiza_Usuario('txtUsuarioAcesso');">
			<%end if%>						
			<%
			if strUsuaAcesso <> "" then
				Response.write strUsuaAcesso
			else
				Response.write Request("hdUsuaAcesso") 
			end if
			%>
		  </td>
		  <td width="35%" height="25"><table width="32%"  border="0">
			<tr>
			  <td><div align="center" class="campo">A&ccedil;&atilde;o</div></td>
			</tr>
			<tr>
			  <td bgcolor="#EFEFEF"><div align="center"><span class="campob"><%=str_Texto_Acao%></span></div></td>
			</tr>
		  </table></td>
		</tr>			
		<tr> 
		  <td width="2%" height="45"></td>
		  <td width="24%" height="45" valign="top" class="campob">Categoria</font>:</td>
		  <td width="39%" height="45">
			<table width="138%" border="0" cellspacing="0" cellpadding="0">
			  <tr> 
				<td width="97%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
				  <input type="radio" name="rdbCategoria" value="A" <%=strChecked_A%>>
				  &quot;A&quot; - Inclus&atilde;o de Detalhamento, Usu&aacute;rio e Consulta.</font></td>
				<td width="3%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;
				</font></td>
			  </tr>
			  <tr>
				<td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
				  <input type="radio" name="rdbCategoria" value="B" <%=strChecked_B%>>
				  <font size="2" face="Verdana, Arial, Helvetica, sans-serif">&quot;B&quot; - Inclus&atilde;o de Detalhamento e Consulta.</font></font>
				  </td>
				<td>&nbsp;</td>
			  </tr>
			  <tr> 
				<td width="97%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
				  <input type="radio" name="rdbCategoria" value="D" <%=strChecked_C%>> 
				  &quot;D&quot; - Consulta.</font></td>
				<td width="3%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
				  </font></td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>		
	<table width="98%" border="0">
	  <tr> 
		<td height="25" colspan="3"> <table width="89%" border="0">
			<tr> 
			  <td width="7%">&nbsp;</td>
			  <td width="93%" class="campob">Acesso de Usuário</td>
			</tr>
		  </table></td>
		<td width="33%">&nbsp;</td>
		<td width="33%">&nbsp;</td>
	  </tr>
	  <tr>
		<td valign="top" class="campo">&nbsp;</td>
		<td valign="top" class="campo">&nbsp;</td>
		<td colspan="3"><table width="100%"  cellpadding="0" cellspacing="0" border="0">
		  <tr>
			<td width="39%" class="campo">Ondas Existentes:</td>
			<td width="5%">&nbsp;</td>
			<td width="40%" class="campo">Ondas Selecionadas: </td>
			<td width="16%">&nbsp;</td>
		  </tr>
		</table></td>
	  </tr>		  
	  <%
	   sqlOnda = "SELECT * FROM ONDA ORDER BY ONDA_TX_DESC_ONDA"
	   Set rds_Onda 	= db_Cogest.Execute(sqlOnda)
	   Set rds_OndaSel 	= db_Cogest.Execute(sqlOnda)
	  %>		  
	  <tr> 
		<td width="1%" valign="top" class="campo">&nbsp;</td>
		<td width="6%" valign="top" class="campo"> <div align="right"> </div></td>
		<td colspan="3"><table width="798" border="0">
			<tr> 
			  <td width="350"> 
				<select name="lstOnda" multiple size="5" class="listResponsavel">				
				  <%
				  if not rds_Onda.bof and not rds_Onda.eof then
					  rds_Onda.movefirst
						Do While not rds_Onda.Eof %>
							<option value="<%=rds_Onda("ONDA_CD_ONDA")%>" ><%=rds_Onda("ONDA_TX_DESC_ONDA")%></option>
						   <% rds_Onda.movenext
						 Loop
				  end if
				  %>
				</select>
			  </td>
			  <td width="32"><table width="30" border="0">
				  <tr> 
					<td width="24"><img src="../img/000030_1.gif" alt="Seleciona Onda" name="imgSetaDireita1" width="20" height="20" id="imgSetaDireita1" onClick="move(document.frm_Usuario.lstOnda,document.frm_Usuario.lstOndaSel,1)" onMouseOver="mOvr(this,'../../../imagens/continua_F02.gif');" onMouseOut="mOut(this,'../../../imagens/continua_F01.gif');"></td>
				  </tr>
				</table></td>
			  <td width="354">				
				<select name="lstOndaSel" size="5" class="listResponsavel" multiple>			
				   <%	 
				   if not rds_OndaSel.bof then
					  rds_OndaSel.movefirst
					  i = 0					  
					  vetAcessoExistentes = split(strAcessoExistentes,",")						  
					  Do While not rds_OndaSel.Eof
						for i = lbound(vetAcessoExistentes) to ubound(vetAcessoExistentes)
							if trim(vetAcessoExistentes(i)) = trim(rds_OndaSel("ONDA_CD_ONDA")) then
							%>						  
								<option value=<%=rds_OndaSel("ONDA_CD_ONDA")%>><%=rds_OndaSel("ONDA_TX_DESC_ONDA")%></option>
							<%
							end if
						next
						rds_OndaSel.movenext							
					  Loop				  
				   end if					  
				   %>		
				</select>
			  </td>
			  <td width="354">&nbsp;</td>
			</tr>
		  </table></td>
	  </tr>
	</table>
	<%
	rds_OndaSel.close
	set rds_OndaSel = nothing	
	
	rds_Onda.close
	set rds_Onda = nothing		
	%>
	
	<table width="625" border="0" align="center">
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td></td>
		  <td></td>
		  <td></td>
		  <td></td>
		  <td>&nbsp;</td>
		  <td><div align="center" class="campo"></div></td>
		</tr>
		<tr>
		  <td width="85" height="37"><a href="javascript:confirma_usuario()"><img src="../img/enviar_01.gif" width="85" height="19" border="0"></a></td>
		  <td width="23" height="37"><b></b></td>
		  <td width="100" height="37"></td>
		  <td width="24" height="37">&nbsp;</td>
		  <td width="90" height="37">
			<%if str_Acao = "A" then%>
				<a href="javascript:confirma_exclusao();"><img src="../img/botao_excluir.gif" border="0"></a>
			<%end if%>	  </td>	  	
		  <td width="10" height="37"></td>
		  <td width="78" height="37"><table width="30" border="0">
            <tr>
              <td width="24"><img src="../img/botao_deletar_on_03.gif" alt="Apaga Onda" name="imgSetaDireita1" width="20" height="20" id="imgSetaDireita1" onClick="deleta(document.frm_Usuario.lstOndaSel,1)" onMouseOver="mOvr(this,'../../../imagens/continua_F02.gif');" onMouseOut="mOut(this,'../../../imagens/continua_F01.gif');"></td>
            </tr>
          </table></td>
		  <td width="104" height="37"><div align="center"></div></td>
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
<!-- InstanceEndEditable -->
    <table width="200" border="0" align="center">
<tr>	
	<td height="10" width="780"></td>
</tr>
<tr>
	<td width="780">			
		<p width="780" align="center"><img src="../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>
