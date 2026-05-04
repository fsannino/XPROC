 
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("txtAcao") <> "0" then
   str_Acao = request("txtAcao")
else
   str_Acao = ""
end if

if request("selMacroPerfil") <> 0 then
   str_MacroPerfil = request("selMacroPerfil")
else
   str_MacroPerfil = "0"
end if

if request("txtNomeTecnico") <> "" then
   str_NomeTecnico = UCase(Trim(request("txtNomeTecnico")))
else
   str_NomeTecnico = ""
end if

if str_NomeTecnico <> "" then
   str_SQL = ""
   str_SQL = str_SQL & " SELECT MCPR_NR_SEQ_MACRO_PERFIL, MEPR_CD_MEGA_PROCESSO,MCPE_TX_SITUACAO "
   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_PERFIL "
   str_SQL = str_SQL & " WHERE MCPE_TX_NOME_TECNICO = '" & str_NomeTecnico & "'"
   'response.write str_SQL      
   set rdsExiste=db.execute(str_SQL)   
   if rdsExiste.EOF then
      rdsExiste.close
      set rdsExiste = Nothing
      response.redirect "msg_ja_existe.asp?opt=3&txtTitFuncao=" & str_Desc_Macro_Perfil
   else
      str_MacroPerfil = rdsExiste("MCPR_NR_SEQ_MACRO_PERFIL")
      st_atual = rdsExiste("MCPE_TX_SITUACAO")
	  str_MegaProcesso = rdsExiste("MEPR_CD_MEGA_PROCESSO")
   end if
   rdsExiste.close
   set rdsExiste = Nothing	  
end if

if str_MacroPerfil <> "" then
   str_SQL = ""
   str_SQL = str_SQL & " SELECT MCPR_NR_SEQ_MACRO_PERFIL, MEPR_CD_MEGA_PROCESSO,MCPE_TX_SITUACAO "
   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_PERFIL "
   str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = '" & str_MacroPerfil & "'"
   'response.write str_SQL      
   set rdsExiste=db.execute(str_SQL)   
   if rdsExiste.EOF then
      rdsExiste.close
      set rdsExiste = Nothing
      response.redirect "msg_ja_existe.asp?opt=3&txtTitFuncao=" & str_Desc_Macro_Perfil
   else
      str_MacroPerfil = rdsExiste("MCPR_NR_SEQ_MACRO_PERFIL")
      st_atual = rdsExiste("MCPE_TX_SITUACAO")
	  str_MegaProcesso = rdsExiste("MEPR_CD_MEGA_PROCESSO")
   end if
   rdsExiste.close
   set rdsExiste = Nothing	  
end if


a=1
if a <> 1 then

   call Exclui_Fisicamente
   a = 1
   if a = 2 then
   str_TP_Del = "Física"
   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MICRO_PERFIL_R3 " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_PERFIL_FUN_NEG " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)    

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MAC_PER_TRAN_OBJETO " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)    

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_PERFIL " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   
   end if
else
   str_SQL = ""
   str_SQL = str_SQL & " SELECT MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO "
   str_SQL = str_SQL & " FROM MACRO_HISTORICO_VALIDACAO "
   str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   str_SQL = str_SQL & " AND MHVA_TX_SITUACAO_MACRO = 'CR'"
   'response.Write(str_SQL)
   set rdsMacroCriado = db.Execute(str_SQL)
   if rdsMacroCriado.EOF then
      ls_teste = " não foi criado "
      call Exclui_Fisicamente
   else	  
      ls_teste = " foi criado "
	   str_TP_Del = "Lógica"
	   if st_atual="CR" or st_atual="AP" then
		  str_Situacao = "ER"
		  str_SQl = ""
		  str_SQl = str_SQL & " Update " & Session("PREFIXO") & "MACRO_PERFIL set " 
		  str_SQl = str_SQL & " MCPE_TX_SITUACAO = '" & str_Situacao & "'"
		  str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
		  str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		  str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"
		  str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
		  db.execute(str_SQl)
	   else
		  set nota=db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil & " AND MHVA_TX_SITUACAO_MACRO='CR'")
		  if nota.eof=false then
			 str_Situacao = "ER"
			 COMENTARIO = "EXCLUÍDO-CRIADO NO R/3-SERÁ EXCLUIDO"		 
			 str_SQl = ""
			 str_SQl = str_SQL & " Update " & Session("PREFIXO") & "MACRO_PERFIL set " 
			 str_SQl = str_SQL & " MCPE_TX_SITUACAO = '" & str_Situacao & "'"
			 str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
			 str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
			 str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"		 
			 str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
			 db.execute(str_SQl) 		 
		  else
			 str_Situacao = "EL"
			 COMENTARIO = "EXCLUÍDO-NÃO CRIADO NO R/3"   		 
			 str_SQl = ""
			 str_SQl = str_SQL & " Update " & Session("PREFIXO") & "MACRO_PERFIL set " 
			 str_SQl = str_SQL & " MCPE_TX_SITUACAO = '" & str_Situacao & "'"
			 str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
			 str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
			 str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"		 
			 str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
			 db.execute(str_SQl) 
		  end if
	   end if
	   str_SQL = "SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & str_MacroPerfil
	   'response.Write(str_SQL)
	   SET HIST = db.EXECUTE(str_SQL)        		
	   ATUAL=HIST("CODIGO")
	   ATUAL = ATUAL + 1         		
	   if atual > 1 then
		  atual = atual
	   else
		  atual=1
	   end if        		
	   SSQL=""
	   SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MEPR_CD_MEGA_PROCESSO, MHVA_TX_COMENTARIO, MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	   SSQL=SSQL+"VALUES(" & str_MegaProcesso & ",'" & COMENTARIO & "', " & ATUAL &", " & str_MacroPerfil & ", '"& str_Situacao &"', 'I', '" & Session("CdUsuario") & "', GETDATE())"      		   
	   'response.Write(SSQL)
	   db.execute(ssql)
	   HIST.CLOSE    
	   str_SQl = "SELECT * FROM MICRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL =" & str_MacroPerfil 
	   'response.Write(str_SQl)
	   set temp=db.execute(str_SQl)
	   if not temp.eof then
		  do while not temp.eof
			 if temp("MICR_TX_SITUACAO") = "CR" or temp("MICR_TX_SITUACAO") = "AP" then
				str_Situacao = "ER"
				COMENTARIO = "EXCLUÍDO-CRIADO NO R/3-SERÁ EXCLUIDO"
				str_SQL = " UPDATE " & Session("PREFIXO") & "MICRO_PERFIL SET "
				str_SQl = str_SQL & " MICR_TX_SITUACAO='" & str_Situacao  & "'"
				str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
				str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
				str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"		 			
				str_SQl = str_SQL & "  WHERE AND MICR_TX_SEQ_MICRO_PERFIL=" & temp("MICR_TX_SEQ_MICRO_PERFIL")
				db.execute(str_SQL)
			 else
				set nota=db.execute("SELECT * FROM " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO WHERE MICR_TX_SEQ_MICRO_PERFIL = " & temp("MICR_TX_SEQ_MICRO_PERFIL") & " AND MHVA_TX_SITUACAO_MACRO='CR'")		 
				if not nota.eof then
				   str_Situacao = "ER"
				   COMENTARIO = "EXCLUÍDO-CRIADO NO R/3-SERÁ EXCLUIDO"
				   str_SQL = " UPDATE " & Session("PREFIXO") & "MICRO_PERFIL SET "
				   str_SQl = str_SQL & " MICR_TX_SITUACAO='" & str_Situacao & "'"
				   str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
				   str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
				   str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"		 						   
				   str_SQl = str_SQL & " WHERE AND MICR_TX_SEQ_MICRO_PERFIL=" & temp("MICR_TX_SEQ_MICRO_PERFIL")
				   db.execute 
				else
				   str_Situacao = "EL"
				   COMENTARIO = "EXCLUÍDO-NÃO CRIADO NO R/3"
				   str_SQL = "UPDATE " & Session("PREFIXO") & "MICRO_PERFIL SET " 
				   str_SQl = str_SQL & " MICR_TX_SITUACAO='" & str_Situacao & "'"
				   str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
				   str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
				   str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"			   
				   str_SQl = str_SQL & " WHERE AND MICR_TX_SEQ_MICRO_PERFIL=" & temp("MICR_TX_SEQ_MICRO_PERFIL")			
				   db.execute 
				end if   
			 end if
			 SET HIST = db.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO WHERE MICR_TX_SEQ_MICRO_PERFIL=" & temp("MICR_TX_SEQ_MICRO_PERFIL"))        		
			 atual=HIST("CODIGO")
			 ATUAL = ATUAL + 1        		
			 if atual > 1 then
				atual = atual
			 else
				atual=1
			 end if
			 HIST.close		 	
			 SSQL=""
			 SSQL="INSERT INTO " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO(MHVA_NR_SEQUENCIA_HIST, MICR_TX_SEQ_MICRO_PERFIL, MHVA_TX_SITUACAO_MICRO,MHVA_TX_COMENTARIO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
			 SSQL=SSQL+"VALUES(" & ATUAL &", '" & temp("MICR_TX_SEQ_MICRO_PERFIL") & "', '"& str_Situacao &"','" & COMENTARIO & "','I', '" & Session("CdUsuario") & "', GETDATE())"
			 db.execute(ssql)
			 temp.movenext
		  Loop		
	   end if
   end if
end if

Sub Exclui_Fisicamente

   str_TP_Del = "Física"
   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MICRO_PERFIL_R3 " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_PERFIL_FUN_NEG " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)    

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MAC_PER_TRAN_OBJETO " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)    

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_PERFIL " 
   str_SQl = str_SQL & " Where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   db.execute(str_SQl)   

end sub

db.Close
set db = Nothing
'response.Write(ls_teste)
response.redirect "msg_ja_existe.asp?opt=2&tpDel=" & str_TP_Del

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">


</script>

<script language="javascript" src="js/troca_lista.js"></script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="../Curso/valida_cad_curso.asp" name="frm1">
        <input type="hidden" name="txtImp" size="20"><input type="hidden" name="txtQua" size="20">
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
      <td colspan="3" height="20">&nbsp;</td>
  </tr>
</table>
        
  <table width="847" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="845">
      </td>
    </tr>
    <tr>
      <td width="845">
        <div align="center"><font face="Verdana" color="#330099" size="3">ser&aacute; 
          redirecionado para outra tela</font></div>
      </td>
    </tr>
    <tr>
      <td width="845">&nbsp;</td>
    </tr>
  </table>
  </form>

</body>

</html>
