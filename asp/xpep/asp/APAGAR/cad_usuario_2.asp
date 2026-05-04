<html>
	<head>		
		<title>Untitled Document</title>
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">	
		<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
		
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
				
				if (document.frm_Usuario.lstMegaProcessoSel.options.length == 0)
				 { 
					 alert("É necessária a seleção de pelo menos uma Onda!");
					 document.frm_Usuario.lstMegaProcessoSel.focus();
					 return;
				 }				
				 else
				 {
				  carrega_txt(document.frm_Usuario.lstMegaProcessoSel);				  
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
	</head>	
	<%
	set db_Cogest = Server.CreateObject("ADODB.Connection")
	db_Cogest.Open Session("Conn_String_Cogest_Gravacao")
			
	str_Acao = "I"
			
	Public Function RetornaNomeUsuario(strChave)		
		sql_VerUsuarioSin= ""			
		sql_VerUsuarioSin = sql_VerUsuarioSin & " SELECT USUA_TX_NOME_USUARIO"		
		sql_VerUsuarioSin = sql_VerUsuarioSin & " FROM XPEP_EQUIPE_SINERGIA "
		sql_VerUsuarioSin = sql_VerUsuarioSin & " WHERE USUA_TX_CD_USUARIO = '" & strChave & "'"
							
		set rds_VerUsuarioSin = db_Cogest.Execute(sql_VerUsuarioSin)
		
		if not rds_VerUsuarioSin.eof then						
			RetornaNomeUsuario = Ucase(rds_VerUsuarioSin("USUA_TX_NOME_USUARIO"))
		else
			sql_VerUsuarioLeg = ""
			sql_VerUsuarioLeg = sql_VerUsuarioLeg & " SELECT USMA_TX_NOME_USUARIO"		
			sql_VerUsuarioLeg = sql_VerUsuarioLeg & " FROM USUARIO_MAPEAMENTO "			
			sql_VerUsuarioLeg = sql_VerUsuarioLeg & " WHERE USMA_CD_USUARIO = '" & strChave & "'"
			
			set rds_VerUsuarioLeg = db_Cogest.Execute(sql_VerUsuarioLeg)
			
			if not rds_VerUsuarioLeg.eof then						
				RetornaNomeUsuario = Ucase(rds_VerUsuarioLeg("USMA_TX_NOME_USUARIO"))			
			else
				RetornaNomeUsuario = "USUÁRIO NÃO LOCALIZADO."
			end if			
			
			rds_VerUsuarioLeg.close
			set rds_VerUsuarioLeg = nothing		
		end if		
		rds_VerUsuarioSin.close
		set rds_VerUsuarioSin = nothing
	End function
	
	strChaveUsuario 	= Request("pChaveUsua")	
	strCampo 			= Request("pCampo")	
			
	if strCampo = "txtUsuarioAcesso" then
		strUsuaAcesso = " - " & RetornaNomeUsuario(strChaveUsuario)
		strUsuarioAcesso = RetornaNomeUsuario(strChaveUsuario)
	end if				
	
	if strUsuaAcesso <> "" then%>	
		<input type="hidden" value="<%=strUsuaAcesso%>" name="hdUsuaAcesso">
	<%else%>
		<input type="hidden" value="<%=Request("hdUsuaAcesso")%>" name="hdUsuaAcesso">
	<%end if	

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
						strAcessoExistentes =  strAcessoExistentes & "|" & rst_ConsAcesso("ONDA_CD_ONDA") 
					end if
					rst_ConsAcesso.movenext
				loop
			else
				strAcessoExistentes = ""
			end if
		else	
			strCategoria = ""
			strAcessoExistentes = ""
		end if
	else
		 strCategoria = ""
	end if
	
	if str_Acao = "I" then
		str_Texto_Acao = "Inclusão"
	elseif str_Acao = "A" then
		str_Texto_Acao = "Alteração"
	end if	
	%>			

	<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
		<form name="frm_Usuario" method="POST">
		  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
			<tr> 
			  <td width="150" height="66" colspan="2">&nbsp;</td>
			  <td width="341" height="66" colspan="2">&nbsp;</td>
			  <td width="276" valign="top" colspan="2" height="66"> 
				<table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
				  <tr> 
					<td bgcolor="#330099" width="39" valign="middle" align="center"> 
					  <div align="center"> 
						<p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
					  </div>
					</td>
					<td bgcolor="#330099" width="36" valign="middle" align="center"> 
					  <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
					</td>
					<td bgcolor="#330099" width="27" valign="middle" align="center"> 
					  <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
					</td>
				  </tr>
				  <tr> 
					<td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
					  <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
					</td>
					<td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
					  <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
					</td>
					<td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
					  <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
					</td>
				  </tr>
				</table>
			  </td>
			</tr>
			<tr bgcolor="#00FF99"> 
			  <td height="20" width="6%">&nbsp; </td>
			  <td height="20" width="3%"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
			  <td height="20" width="21%"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Enviar</b></font></td>
			  <td colspan="2" height="20"> 
				<div align="right"><a href="javascript:Limpa()"><img src="../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></div>
			  </td>
			  <td height="20" width="39%"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Limpa</b></font></td>
			</tr>
		  </table>
		  <table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr> 
			  <td width="24%">&nbsp;</td>
			  <td width="62%">&nbsp;</td>
			  <td width="14%">&nbsp;</td>
			</tr>
			<tr> 
			  <td width="24%">&nbsp;</td>
			  <td width="62%">Cadastro de Usu&aacute;rio</td>
			  <td width="14%">&nbsp;</td>
			</tr>
		  </table>
		  <table width="101%" border="0" cellspacing="0" cellpadding="0" height="148">
			<tr> 
			  <td width="2%" height="21">&nbsp;</td>
			  <td width="24%" height="21">&nbsp;</td>
			  <td width="37%" height="21">&nbsp;</td>
			  <td width="37%" height="21">&nbsp;</td>
			</tr>				
			<tr> 
			  <td width="2%" height="25"></td>
			  <td width="24%" height="25" class="campob">Chave do Usu&aacute;rio:</td>
			  <td width="37%" height="25"> 	
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
			  <td width="37%" height="25"><table width="32%"  border="0">
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
			  <td width="37%" height="45">
				<table width="110%" border="0" cellspacing="0" cellpadding="0">
				  <tr> 
					<td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
					  <input type="radio" name="rdbCategoria" value="A" <%=strChecked_A%>>
					  &quot;A&quot;</font></td>
					<td width="30%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;
				    </font></td>
				  </tr>
				  <tr>
				    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
				      <input type="radio" name="rdbCategoria" value="B" <%=strChecked_B%>>
                      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">&quot;B&quot;</font></font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;
			          </font></td>
				    <td>&nbsp;</td>
			      </tr>
				  <tr> 
					<td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
					  <input type="radio" name="rdbCategoria" value="C" <%=strChecked_C%>> 
					  &quot;C&quot;</font></td>
					<td width="30%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
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
		   sqlMegaProcesso = "SELECT * FROM MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO"
		   Set rds_MegaProcesso 	= db_Cogest.Execute(sqlMegaProcesso)
		   Set rds_MegaProcessoSel 	= db_Cogest.Execute(sqlMegaProcesso)
		  %>		  
		  <tr> 
			<td width="1%" valign="top" class="campo">&nbsp;</td>
			<td width="6%" valign="top" class="campo"> <div align="right"> </div></td>
			<td colspan="3"><table width="798" border="0">
				<tr> 
				  <td width="350"> 
					<select name="lstMegaProcesso" multiple size="5" class="listResponsavel">				
					  <%
					  if not rds_MegaProcesso.bof and not rds_MegaProcesso.eof then
						  rds_MegaProcesso.movefirst
							Do While not rds_MegaProcesso.Eof %>
							    <option value="<%=rds_MegaProcesso("MEPR_CD_MEGA_PROCESSO")%>" ><%=rds_MegaProcesso("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
							   <% rds_MegaProcesso.movenext
							 Loop
					  end if
					  %>
					</select>
				  </td>
				  <td width="32"><table width="30" border="0">
					  <tr> 
						<td width="24"><img src="../img/000030_1.gif" alt="Seleciona Onda" name="imgSetaDireita1" width="20" height="20" id="imgSetaDireita1" onClick="move(document.frm_Usuario.lstMegaProcesso,document.frm_Usuario.lstMegaProcessoSel,1)" onMouseOver="mOvr(this,'../../../imagens/continua_F02.gif');" onMouseOut="mOut(this,'../../../imagens/continua_F01.gif');"></td>
					  </tr>
					</table></td>
				  <td width="354">	  
					<select name="lstMegaProcessoSel" size="5" class="listResponsavel" multiple>			
					   <%	 
					   if not rds_MegaProcessoSel.bof then
						  rds_MegaProcessoSel.movefirst
						  i = 0
						  vetAcessoExistentes = split(strAcessoExistentes,"|")						  
						  Do While not rds_MegaProcessoSel.Eof
							for i = lbound(vetAcessoExistentes) to ubound(vetAcessoExistentes)
								if trim(vetAcessoExistentes(i)) = trim(rds_MegaProcessoSel("MEPR_CD_MEGA_PROCESSO")) then
								%>						  
									<option value=<%=rds_MegaProcessoSel("MEPR_CD_MEGA_PROCESSO")%>><%=rds_MegaProcessoSel("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
								<%
								end if
							next
							rds_MegaProcessoSel.movenext							
						  Loop				  
					   end if					  
					   %>		
					</select>
				  </td>
				  <td width="354"><table width="30" border="0">
					<tr>
					  <td width="24"><img src="../img/botao_deletar_on_03.gif" alt="Apaga Onda" name="imgSetaDireita1" width="20" height="20" id="imgSetaDireita1" onClick="deleta(document.frm_Usuario.lstMegaProcessoSel,1)" onMouseOver="mOvr(this,'../../../imagens/continua_F02.gif');" onMouseOut="mOut(this,'../../../imagens/continua_F01.gif');"></td>
					</tr>
				  </table></td>
				</tr>
			  </table></td>
		  </tr>
		</table>
		<%
		rds_MegaProcesso.close
		set rds_MegaProcesso = nothing		
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
			  <td width="78" height="37">&nbsp;</td>
			  <td width="104" height="37"><div align="center"></div></td>
			</tr>
		  </table>
  		  <input type="hidden" value="<%=str_Acao%>" name="pAcao">		  		 
		  <input type="hidden" value="" name="pChaveUsua">
		  <input type="hidden" value="<%=strUsuarioAcesso%>" name="pNomeUsua">
		  <input type="hidden" name="pOndaSelecionada">		  
		</form>
	</body>
</html>
