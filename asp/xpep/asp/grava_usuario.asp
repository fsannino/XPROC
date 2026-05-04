<%					
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

strAcao 			= Request("pAcao")
strCdUsuario 		= Ucase(Request("txtUsuarioAcesso"))
strNomeUsuario 		= Ucase(Request("pNomeUsua"))
strCategoria 		= Request("rdbCategoria")
strOndaSelecionada 	= Request("pOndaSelecionada")

'response.write "<br><br><br> strAcao " & strAcao & "<br>"
'response.write "strCdUsuario " & strCdUsuario  & "<br>"
'response.write "strNomeUsuario " & strNomeUsuario  & "<br>"
'response.write "strCategoria " & strCategoria  & "<br>"
'response.write "strOndaSelecionada " & strOndaSelecionada   & "<br>"
'response.end

db_Cogest.execute("DELETE XPEP_ACESSO WHERE USUA_CD_USUARIO='" & strCdUsuario & "'")

if strAcao = "I" then

	'Call VerificaUsuarioExistente("Chave do Usuário",strCdUsuario, "Sinergia")

	on error resume next			
		sqlNovoUsuario = ""
		sqlNovoUsuario = " INSERT INTO XPEP_USUARIO(USUA_CD_USUARIO, USUA_TX_NOME_USUARIO, "
		sqlNovoUsuario = sqlNovoUsuario & "USUA_TX_CATEGORIA, ATUA_TX_OPERACAO, "
		sqlNovoUsuario = sqlNovoUsuario & "ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO"
		sqlNovoUsuario = sqlNovoUsuario & ")VALUES('" & strCdUsuario & "','" & strNomeUsuario & "','" & strCategoria & "',"
		sqlNovoUsuario = sqlNovoUsuario & "'I','" & Session("CdUsuario") & "',GETDATE())"
		'Response.write sqlNovoUsuario & "<br><br>"
		'Response.end							
		db_Cogest.Execute(sqlNovoUsuario)
			
		countAcesso	= 0
		vetOndaSelecionada = split(strOndaSelecionada,",")
		for countAcesso = lbound(vetOndaSelecionada) to ubound(vetOndaSelecionada)			
			if vetOndaSelecionada(countAcesso) <> "" then
				sql_NovoAcesso = ""
				sql_NovoAcesso = sql_NovoAcesso & " INSERT INTO XPEP_ACESSO(USUA_CD_USUARIO, ONDA_CD_ONDA, "
				sql_NovoAcesso = sql_NovoAcesso & " ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, "
				sql_NovoAcesso = sql_NovoAcesso & " ATUA_DT_ATUALIZACAO"
				sql_NovoAcesso = sql_NovoAcesso & ")VALUES('" & strCdUsuario & "'," & cint(vetOndaSelecionada(countAcesso)) & ","
				sql_NovoAcesso = sql_NovoAcesso & "'I','" & Session("CdUsuario") & "',GETDATE())"	
				'Response.write sql_NovoAcesso & "<br><br>"
				'Response.end			
				db_Cogest.Execute(sql_NovoAcesso)	
			end if
		next			
			
		if err.number = 0 then		
			strMSG = "Usuário incluído com sucesso."
		else
			strMSG = "Houve um erro na inclusão de Usuário."
		end if		
		
elseif strAcao = "A" then
	
	on error resume next			
		sqlAltUsuario = ""
		sqlAltUsuario = "UPDATE XPEP_USUARIO SET"
		sqlAltUsuario = sqlAltUsuario & " USUA_TX_NOME_USUARIO ='" & strNomeUsuario & "'"
		sqlAltUsuario = sqlAltUsuario & ",USUA_TX_CATEGORIA ='" & strCategoria & "'"
		sqlAltUsuario = sqlAltUsuario & ",ATUA_TX_OPERACAO ='A'" 
		sqlAltUsuario = sqlAltUsuario & ",ATUA_CD_NR_USUARIO ='" & Session("CdUsuario") & "'"
		sqlAltUsuario = sqlAltUsuario & ",ATUA_DT_ATUALIZACAO = GETDATE()"
		sqlAltUsuario = sqlAltUsuario & " WHERE USUA_CD_USUARIO = '" & strCdUsuario & "'"			
		'Response.write sqlAltUsuario & "<br><br>"
		'Response.end							
		db_Cogest.Execute(sqlAltUsuario)
			
		countAcesso	= 0
		
		vetOndaSelecionada = split(strOndaSelecionada,",")
		for countAcesso = lbound(vetOndaSelecionada) to ubound(vetOndaSelecionada)			
			if vetOndaSelecionada(countAcesso) <> "" then				
				sql_NovoAcesso = ""
				sql_NovoAcesso = sql_NovoAcesso & " INSERT INTO XPEP_ACESSO(USUA_CD_USUARIO, ONDA_CD_ONDA, "
				sql_NovoAcesso = sql_NovoAcesso & " ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, "
				sql_NovoAcesso = sql_NovoAcesso & " ATUA_DT_ATUALIZACAO"
				sql_NovoAcesso = sql_NovoAcesso & ")VALUES('" & strCdUsuario & "'," & cint(vetOndaSelecionada(countAcesso)) & ","
				sql_NovoAcesso = sql_NovoAcesso & "'I','" & Session("CdUsuario") & "',GETDATE())"	
				'Response.write sql_NovoAcesso & "<br><br>"
				'Response.end			
				db_Cogest.Execute(sql_NovoAcesso)	
			end if
		next			
			
		if err.number = 0 then		
			strMSG = "Usuário alterado com sucesso."
		else
			strMSG = "Houve um erro na alteração do Usuário."
		end if	
		
elseif strAcao = "E" then
			
	sqlExcUsuario = ""
	sqlExcUsuario = sqlExcUsuario & " DELETE XPEP_USUARIO"		
	sqlExcUsuario = sqlExcUsuario & " WHERE USUA_CD_USUARIO = '" & strCdUsuario & "'"		
	'Response.write sqlExcUsuario & "<br><br>"
	'Response.end															
																							
	on error resume next
		db_Cogest.Execute(sqlExcUsuario)	
			
	if err.number = 0 then		
		strMSG = "Usuário excluído com sucesso."
	else
		strMSG = "Houve um erro na exclusão do Usuário."
	end if		
end if

'Public Function VerificaUsuarioExistente(strCampo, strChave, strTipoResponsavel) 
'		
'	sql_VerUsuario= ""	
'	if strTipoResponsavel = "Sinergia" then 
'		sql_VerUsuario = sql_VerUsuario & " SELECT USUA_CD_USUARIO"		
'		sql_VerUsuario = sql_VerUsuario & " FROM USUARIO "
'		sql_VerUsuario = sql_VerUsuario & " WHERE USUA_CD_USUARIO = '" & strChave & "'"
'	elseif strTipoResponsavel = "Legado" then				
'		sql_VerUsuario = sql_VerUsuario & " SELECT USMA_CD_USUARIO"		
'		sql_VerUsuario = sql_VerUsuario & " FROM USUARIO_MAPEAMENTO "
'		sql_VerUsuario = sql_VerUsuario & " WHERE USMA_TX_MATRICULA <> 0"
'		sql_VerUsuario = sql_VerUsuario & " AND USMA_CD_USUARIO = '" & strChave & "'"
'	end if
	
'	set rds_VerUsuario = db_Cogest.Execute(sql_VerUsuario)
	
'	if rds_VerUsuario.eof then				
'		strMsg = "Favor verificar a chave informada (" & strChave & "). No campo " & strCampo & "!"		
'		rds_VerUsuario.close
'		set rds_VerUsuario = nothing		
'		Response.redirect "msg_geral.asp?pMsg=" & strMsg & "&pPlano=&pUsua=OK"
'	end if
'	rds_VerUsuario.close
'	set rds_VerUsuario = nothing
'End function
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
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
<script language="JavaScript">	

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
	<table width="849" height="195" border="0" cellpadding="5" cellspacing="5">
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
				
		  <td width="117" height="29"></td>
				
		  <td width="53" height="29" valign="middle" align="left"></td>
				
		  <td height="29" valign="middle" align="left" colspan="2"> 
		  <%if err.number=0 then%>
		  <b><font face="Verdana" color="#330099" size="2"><%=strMSG%></font></b> 
		  </td>				
			  </tr>
		  <%else%>    
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  <b><font face="Verdana" size="2" color="#800000"><%=strMSG%> - <%=err.description%></font></b> 
		  </td>
			  </tr>
			  <%end if%>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" width="32"> 
			<a href="../../../indexA_xpep.asp">
			<img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>				
		  <td height="1" valign="middle" align="left" width="629"> 
			<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
			  </tr>			  
			  <tr>				
				  <td width="117" height="1"></td>						
				  <td width="53" height="1" valign="middle" align="left"></td>						
				  <td height="1" valign="middle" align="left" width="32">					
					<a href="cad_usuario.asp"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
				  <td height="1" valign="middle" align="left" width="629"> 
					<font face="Verdana" color="#330099" size="2">Retornar para Tela de Usuário</font>
				  </td>
			  </tr>		
			  <tr>					
			  <td width="117" height="1"></td>					
			  <td width="53" height="1" valign="middle" align="left"></td>					
			  <td height="1" valign="middle" align="left" colspan="2"> 
			  </td>
			  </tr>
			</table>
  <table width="614" border="0">
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>		
  </table>
  	<%
	db_Cogest.close
	set db_Cogest = nothing
	%>
  <p>&nbsp;</p>
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
