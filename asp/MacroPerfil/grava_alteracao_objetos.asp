<%@LANGUAGE="VBSCRIPT"%> 
 
<%
str_Opt = Request("txtOPT")

if request("txtFuncao") <> "0" then
   str_Funcao = request("txtFuncao")
else
   str_Funcao = 0
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_MacroPerfil=request("txtMacroPerfil")
str_Transacao=request("txtTransacao")
int_QtdObj=CInt(request("txtQtdObj"))

str_SQl = ""
str_SQl = str_SQl & " Select MCPE_TX_NOME_TECNICO from " & Session("PREFIXO") & "MACRO_PERFIL "
str_SQl = str_SQl & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil

set rsd_MacroPerfil=conn_db.execute(str_SQL)
IF not rsd_MacroPerfil.EOF then
   str_Nome_Tecnico = rsd_MacroPerfil("MCPE_TX_NOME_TECNICO")
else
   str_Nome_Tecnico = "não achou NT"
end if
rsd_MacroPerfil.close
set rsd_MacroPerfil = Nothing
	
int_contador = 1
int_Tot_Gravado = 0
'response.write int_QtdObj
'response.write "   -   "
do while int_contador <= int_QtdObj
   str_Obj = request("txtObj" & int_contador) 
   str_Campo =  request("txtCampo" & int_contador)
   str_Valor =  request("txtValor" & int_contador)
   str_Valor2 =  request("txtValorPadrao" & int_contador)
   'response.write "  Valor 1 : "
   'response.write str_Valor
   'response.write "  Valor 2 : "
   'response.write str_Valor2
   if Trim(str_Valor) <> Trim(str_Valor2) Then
      call f_Altera_Objeto(str_MacroPerfil,str_Transacao,str_Obj,str_Campo,str_Valor)
	  int_Tot_Gravado = int_Tot_Gravado + 1
   end if	  
   int_contador = int_contador + 1
'  response.write int_contador
loop

Sub f_Altera_Objeto(p_MacroPerfil, p_Transacao, p_Obj, p_Campo, p_Valor)
	str_SQL = ""
	str_SQl = str_SQl & " UPDATE " & Session("PREFIXO") & "MAC_PER_TRAN_OBJETO "
	str_SQl = str_SQl & " SET MPTO_TX_VALORES = '" & p_Valor & "'"
	if str_Opt = 1 or str_Opt = 3 then
	   str_SQl = str_SQl & " ,MPTO_TX_SIT_ALTERACAO_VALOR = '1'"
	else
	   str_SQl = str_SQl & " ,MPTO_TX_SIT_ALTERACAO_VALOR1 = '1'"
	end if   
	str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
	str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
	str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"
    str_SQl = str_SQl & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & p_MacroPerfil
    str_SQl = str_SQl & " AND TRAN_CD_TRANSACAO = '" & p_Transacao & "'" 
    str_SQl = str_SQl & " AND TROB_TX_OBJETO = '" & p_Obj & "'" 
    str_SQl = str_SQl & " AND TROB_TX_CAMPO = '" & p_Campo & "'" 
	'response.write str_SQl	
	conn_db.execute(str_SQl)
	
	str_SQl = ""
	str_SQl = str_SQl & " Update " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO "
	if str_Opt = 1 or str_Opt = 3 then
       str_SQl = str_SQl & " SET MCPT_NR_SITUACAO_ALTERACAO = 1 "
	else
       str_SQl = str_SQl & " SET MCPT_NR_SITUACAO_ALTERACAO1 = 1 "
	end if   
	str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
	str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
	str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"	
	str_SQl = str_SQl & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
	str_SQl = str_SQl & " AND TRAN_CD_TRANSACAO = '" & str_Transacao & "'" 
	conn_db.execute(str_SQl)
    
    Set rst_Situacao = conn_db.EXECUTE("SELECT MCPE_TX_SITUACAO FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & str_MacroPerfil)
	str_Situacao = rst_Situacao("MCPE_TX_SITUACAO")
	rst_Situacao.Close
	set rst_Situacao = Nothing
	
	call Verifica_Dono(str_Transacao)
	
	if (str_Opt = 1 or str_Opt = 3) and str_Situacao <> "EE" then
       str_SQl = ""
	   str_SQl = str_SQl & " Update " & Session("PREFIXO") & "MACRO_PERFIL "
       str_SQl = str_SQl & " SET MCPE_TX_SITUACAO = 'EE' "
	   str_SQl = str_SQl & " ,ATUA_TX_OPERACAO = 'A'"   
	   str_SQl = str_SQl & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
	   str_SQl = str_SQl & " ,ATUA_DT_ATUALIZACAO = GETDATE()"	   
	   str_SQl = str_SQl & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
	   conn_db.execute(str_SQl)
	   
	   SET HIST = CONN_DB.EXECUTE("SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & str_MacroPerfil)
        		
       ATUAL=HIST("CODIGO")
       ATUAL = ATUAL + 1
        		
       if atual > 1 then
	   	  atual = atual
   	   else
   	      atual=1
   	   end if

	   SSQL=""
	   SSQL="INSERT INTO " & Session("PREFIXO") & "MACRO_HISTORICO_VALIDACAO(MHVA_NR_SEQUENCIA_HIST, MCPR_NR_SEQ_MACRO_PERFIL, MHVA_TX_SITUACAO_MACRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
	   SSQL=SSQL+"VALUES(" & atual &", " & str_MacroPerfil & ", 'EE', 'I', '" & Session("CdUsuario") & "', GETDATE())"        		
	   conn_db.execute(ssql)
	   'response.write ssql
	   
	end if   	
end sub

' ================================== VERIFICA OS DONOS =================
Sub Verifica_Dono(p_Transacao)
      'int_Qtd_Mega = 0
      'str_ListaMega = ""
      'str_Necessita_Autorizacao = 0
      str_SQL = ""
	  str_SQL = str_SQL & " SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
	  str_SQL = str_SQL & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
      str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO_MEGA INNER JOIN"
      str_SQL = str_SQL & " " & Session("PREFIXO") & "MEGA_PROCESSO ON "
      str_SQL = str_SQL & " " & Session("PREFIXO") & "TRANSACAO_MEGA.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
      str_SQL = str_SQL & " WHERE " & Session("PREFIXO") & "TRANSACAO_MEGA.TRAN_CD_TRANSACAO = '" & p_Transacao & "'" 
	  str_SQL = str_SQL & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "				   
	  Set rdsExiste2 = conn_db.Execute(str_SQL)
	  'response.Write(str_SQL)				   
	  loo_Existe = False
	  IF not rdsExiste2.EOF then
	     Do While not rdsExiste2.EOF
	         if InStr("," & Session("AcessoUsuario") & ",","," &  Trim(rdsExiste2("MEPR_CD_MEGA_PROCESSO")) & ",") <> 0 then						 
	           loo_Existe = True
               exit do
	        end if
			rdsExiste2.Movenext
		 Loop
	  else
	     loo_Existe = True	 
	  end if
	  if loo_Existe = False then 'and not rdsExiste2.eof
         rdsExiste2.MoveFirst
         if not rdsExiste2.EOF then
	        Do While not rdsExiste2.EOF
			   if Ja_Registrado_Autorizacao(str_MacroPerfil, p_Transacao, rdsExiste2("MEPR_CD_MEGA_PROCESSO")) = false then
                   Call Grava_Para_Autorizar(str_MacroPerfil, p_Transacao, rdsExiste2("MEPR_CD_MEGA_PROCESSO"))
			   end if   
			   'str_Necessita_Autorizacao = 1
	           rdsExiste2.Movenext
		    Loop
	      end if		
	  end if
	  rdsExiste2.close
	  set rdsExiste2 = Nothing

end Sub

'   if str_Necessita_Autorizacao = "0" then 
'      str_SQL = ""
'	  str_SQL = str_SQL & " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET " 
'	  str_SQL = str_SQL & " MCPE_TX_SITUACAO = 'EC' "
'	  str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & sequencia
 '     db.execute(str_SQL)
'   end if	  

Function Ja_Registrado_Autorizacao(pMacro, pTransacao, pMega)
	str_SQL = ""
	str_SQL = str_SQL & " SELECT MCPR_NR_SEQ_MACRO_PERFIL, TRAN_CD_TRANSACAO, "
    str_SQL = str_SQL & " MEPR_CD_MEGA_PROCESSO, MAOA_TX_AUTORIZADO"
    str_SQL = str_SQL & " FROM dbo.MACRO_OBJ_AUTORIZA"
    str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & pMacro
    str_SQL = str_SQL & " and TRAN_CD_TRANSACAO = '" & pTransacao & "'"
    str_SQL = str_SQL & " and MEPR_CD_MEGA_PROCESSO = " & pMega
	set rsExisteAuto = conn_db.Execute(str_SQL)
	if rsExisteAuto.EOF then
	   Ja_Registrado_Autorizacao = False
	else
	   Ja_Registrado_Autorizacao = True
	   str_SQL = ""
       str_SQL = str_SQL & " Update dbo.MACRO_OBJ_AUTORIZA"
	   str_SQL = str_SQL & " set MAOA_TX_AUTORIZADO = '0'"
       str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & pMacro
       str_SQL = str_SQL & " and TRAN_CD_TRANSACAO = '" & pTransacao & "'"
       str_SQL = str_SQL & " and MEPR_CD_MEGA_PROCESSO = " & pMega
	   conn_db.execute(str_SQL)
	end if
	
end Function

Sub Grava_Para_Autorizar(pMacro, pTransacao, pMega)
	str_SQL = ""
	str_SQL = str_SQL & " Insert into " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA( "
   str_SQL = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL "
	str_SQL = str_SQL & " , TRAN_CD_TRANSACAO"
	str_SQL = str_SQL & " , MEPR_CD_MEGA_PROCESSO"
	str_SQL = str_SQL & " , MAOA_TX_AUTORIZADO"
	str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
	str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
	str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO "
	str_SQL = str_SQL & " ) values ( "
	str_SQL = str_SQL & pMacro 
	str_SQL = str_SQL & ",'" & pTransacao & "'"
	str_SQL = str_SQL & "," & pMega
	str_SQL = str_SQL & ",'0'" 
    str_SQL = str_SQL & ",'I'"
    str_SQL = str_SQL & ",'" & Session("CdUsuario") & "',GETDATE())"
	'RESPONSE.WRITE STR_SQL
	conn_db.execute(str_SQL)
end sub


%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK REL="SHORTCUT ICON" href="http://regina/imagens/Wrench.ico">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" action="../valida_altera_processo.asp?mega=<%=str_MegaProcesso%>&Proc=<%=str_Processo%>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr>
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="3%">&nbsp;</td>
      <td height="20" width="43%">opt : <%=str_Opt%> </td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="39%">&nbsp; </td>
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
      <td width="62%"> <font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Grava 
        altera&ccedil;&atilde;o de valores de Objetos - Macro Perfil</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_MegaProcesso%></td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <%if erro=0 then%>
    <tr> 
      <td width="3%"></td>
      <td width="24%"><%'=str_Processo%></td>
      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><b>Opera&ccedil;&atilde;o 
        realizada com Sucesso! Total gravado : <%=int_Tot_Gravado%></b></font></td>
      <td width="14%"></td>
    </tr>
    <% else %>
    <tr> 
      <td width="3%"></td>
      <td width="24%">&nbsp;</td>
      <td width="59%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#800000">Não 
        foi possível realizar a opera&ccedil;&atilde;o. Avise o problema.</font></b></td>
      <td width="14%"></td>
    </tr>
    <% end if %>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_SubProcesso%></td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_Cenario%></td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">
        <table width="74%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr> 
            <td height="41"><a href="rel_funcao_transacao.asp?txtMacroPerfil=<%=str_MacroPerfil%>&selFuncao=<%=str_Funcao%>&txtNomeTecnico=<%=str_Nome_Tecnico%>&txtOPT=<%=str_Opt%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela de Edi&ccedil;&atilde;o de Objetos</font></td>
          </tr>
          <tr> 
            <td height="41"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela Principal</font></td>
          </tr>
        </table>
      </td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
</body>
</html>
