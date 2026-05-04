<%
CONECTA="Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"
set db = Server.CreateObject("ADODB.Connection")
db.Open CONECTA

if request("txtAcao") <> "0" then
   str_Acao = request("txtAcao")
else
   str_Acao = ""
end if

if request("txtMacroPerfil") <> 0 then
   str_MacroPerfil = request("txtMacroPerfil")
else
   str_MacroPerfil = ""
end if

if request("selMegaProcesso") <> 0 then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = ""
end if

if request("txtPrefixoNomeTecnico") <> "0" then
   str_PrefixoNomeTecnico = UCase(Trim(request("txtPrefixoNomeTecnico")))
else
   str_PrefixoNomeTecnico = ""
end if

if request("txtNomeTecnico") <> "0" then
   str_NomeTecnico = UCase(Trim(request("txtNomeTecnico")))
else
   str_NomeTecnico = ""
end if

if request("txtDescMacroPerfil") <> "0" then
   str_DescMacroPerfil = Ucase(Trim(request("txtDescMacroPerfil")))
else
   str_DescMacroPerfil = ""
end if

if request("selFuncPrinc") <> "0" then
   str_FuncPrinc = Ucase(Trim(request("selFuncPrinc")))
else
   str_FuncPrinc = ""
end if

str_SQL = ""
str_SQL = str_SQL & " SELECT MCPE_TX_DESC_MACRO_PERFIL  "
str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_PERFIL "
str_SQL = str_SQL & " WHERE MCPE_TX_NOME_TECNICO = '" & ucase(str_PrefixoNomeTecnico) & ucase(str_NomeTecnico) & "'"
if str_Acao = "M" then
   str_SQL = str_SQL & " AND MCPR_NR_SEQ_MACRO_PERFIL <> " & str_MacroPerfil
end if

set rdsExiste=db.execute(str_SQL)
if not rdsExiste.EOF then
   str_Desc_Macro_Perfil = rdsExiste("MCPE_TX_DESC_MACRO_PERFIL")
   rdsExiste.close
   set rdsExiste = Nothing
   response.redirect "msg_ja_existe.asp?opt=0&txtTitFuncao=" & str_Desc_Macro_Perfil
end if
rdsExiste.close
set rdsExiste = Nothing

' "C" opção de CRIACAO
if str_Acao = "C" then
   set temp=db.execute("SELECT MAX(MCPR_NR_SEQ_MACRO_PERFIL)AS CODIGO FROM " & Session("PREFIXO") & "MACRO_PERFIL")

   if not isnull(temp("codigo")) then
      sequencia=temp("CODIGO")+1
   else
      sequencia=1
   end if
   temp.close
   set temp = Nothing
   
   ssql="INSERT INTO " & Session("PREFIXO") & "MACRO_PERFIL  "
   ssql=ssql+ "(MCPR_NR_SEQ_MACRO_PERFIL, MCPE_TX_DESC_MACRO_PERFIL, MCPE_TX_NOME_TECNICO"
   ssql=ssql+ ", MCPE_TX_SITUACAO, MEPR_CD_MEGA_PROCESSO " 
   ssql=ssql+ ", ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO , ATUA_DT_ATUALIZACAO )"
   ssql=ssql+ " VALUES( " & sequencia & ",'" & ucase(str_DescMacroPerfil) & "', "
   ssql=ssql+ "'" & ucase(str_PrefixoNomeTecnico) & ucase(str_NomeTecnico) & "', 'EE', "
   ssql=ssql+ str_MegaProcesso & ", "
   ssql=ssql+ "'I','" & Session("CdUsuario") & "',GETDATE())"
   db.execute(ssql)   
   call Grava_Funcao(sequencia, str_FuncPrinc, 0)
else
   ssql=""
   ssql=ssql+ " UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET  "
   ssql=ssql+ " MCPE_TX_DESC_MACRO_PERFIL = '" & ucase(str_DescMacroPerfil) & "'" 
   ssql=ssql+ " , MCPE_TX_NOME_TECNICO = '" & ucase(str_PrefixoNomeTecnico) & ucase(str_NomeTecnico) & "'"
   ssql=ssql+ " , ATUA_TX_OPERACAO = 'A'"
   ssql=ssql+ " , ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
   ssql=ssql+ " , ATUA_DT_ATUALIZACAO = GETDATE()"
   ssql=ssql+ " where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil

   db.execute(ssql)
   sequencia = str_MacroPerfil
   
   str_SQL = ""
   str_SQL = str_SQL & " Delete from " & Session("PREFIXO") & "MACRO_PERFIL_FUN_NEG "
   str_SQL = str_SQL & " where MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   str_SQL = str_SQL & " AND MPFN_NR_TIPO = 1 "
   'response.write str_SQL
   db.execute(str_SQL)
   
end if

Sub Grava_Funcao(strSeq, strFunc, tipo)
   str_SQL = ""
   str_SQL = str_SQL & " insert into " & Session("PREFIXO") & "MACRO_PERFIL_FUN_NEG "
   str_SQL = str_SQL & " (FUNE_CD_FUNCAO_NEGOCIO "
   str_SQL = str_SQL & " ,MCPR_NR_SEQ_MACRO_PERFIL "
   str_SQL = str_SQL & " ,MPFN_NR_TIPO "
   str_SQL = str_SQL & " , ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO , ATUA_DT_ATUALIZACAO)"
   str_SQL = str_SQL & " VALUES( '" & strFunc & "'," & strSeq & ", " & tipo & ","
   str_SQL = str_SQL & "'I','" & Session("CdUsuario") & "',GETDATE())"
   'response.write str_SQL
   set temp=db.execute(str_SQL)
 ' temp.close
 ' set temp = Nothing
end sub

str_valor=request("txtFuncSelec")

'Coloca uma virgula no fim de string, se não houver
if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if

'Pega o tamanho da string
tamanho = Len(str_valor)

'Retira a vírgula do início da string, se houver
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If
'response.write str_valor
'Atualiza o tamanho da string
tamanho = Len(str_valor)

'Inicializa o Contador
contador = 1

'início da Rotina
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)  
	'	response.write " sequencia : "
'		response.write sequencia
'		response.write " atual:  "
 '       response.write str_atual
        if str_atual <> str_FuncPrinc then
           call Grava_Funcao(sequencia, str_atual, 1)
		end if   
	    valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop

if str_Acao = "C" then
   str_SQL = ""
   str_SQL = str_SQL & " INSERT INTO " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO( "
   str_SQL = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL, TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO "
   str_SQL = str_SQL & " , MCPT_NR_SITUACAO_ALTERACAO, MCPT_NR_SITUACAO_ALTERACAO_FUNC "
   str_SQL = str_SQL & " , ATUA_TX_OPERACAO "
   str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO ) "
   str_SQL = str_SQL & " (SELECT DISTINCT " & sequencia & ", TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO, "
   str_SQL = str_SQL & " 0, 0, "
   str_SQL = str_SQL & " 'I', '" & Session("CdUsuario") &  "', GETDATE() "
   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
   str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & str_FuncPrinc & "') "

   db.execute(str_SQL)
   
   str_SQL = ""
   str_SQL = str_SQL & " Select distinct MEPR_CD_MEGA_PROCESSO "
   str_SQL = str_SQL & " FROM MACRO_PERFIL_TRANSACAO "
   str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & pMacro
   str_SQL = str_SQL & " AND MEPR_CD_MEGA_PROCESSO <> " str_MegaProcesso	
   set rs_QtdMega=conn_db.execute(str_SQL)
   int_Qtd_Mega = 0
   str_ListaMega = ""
   Do while not rs_QtdMega
      int_Qtd_Mega = int_Qtd_Mega + 1
	  str_ListaMega = str_ListaMega & rs_QtdMega("MEPR_CD_MEGA_PROCESSO") & ","
	  rs_QtdMega.movenext	   
   loop
   if int_Qtd_Mega > 0 then
	  rs_QtdMega.MoveFirst
	  Do while not rs_QtdMega
         int_Qtd_Mega = int_Qtd_Mega + 1
         str_SQL = ""
		 str_SQL = str_SQL & " Insert  
		 str_SQL = str_SQL & " 
	     rs_QtdMega.movenext	   
	  loop


   str_SQL = ""
   str_SQL = str_SQL & " INSERT INTO " & Session("PREFIXO") & "MAC_PER_TRAN_OBJETO( "
   str_SQL = str_SQL & " MCPR_NR_SEQ_MACRO_PERFIL, TRAN_CD_TRANSACAO, "
   str_SQL = str_SQL & " TROB_TX_CAMPO, TROB_TX_OBJETO, MPTO_TX_VALORES, TROB_TX_CRITICO , ATUA_TX_OPERACAO, "
   str_SQL = str_SQL & " ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO ) "
   str_SQL = str_SQL & " (SELECT " & sequencia & ", TRAN_CD_TRANSACAO, TROB_TX_CAMPO, "
   str_SQL = str_SQL & " TROB_TX_OBJETO, TRON_TX_VALORES, TROB_TX_CRITICO, "
   str_SQL = str_SQL & " 'I', '" & Session("CdUsuario") &  "', GETDATE() "
   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO_OBJETO "
   str_SQL = str_SQL & " WHERE TRAN_CD_TRANSACAO IN (SELECT TRAN_CD_TRANSACAO "
   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
   str_SQL = str_SQL & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & str_FuncPrinc & "')) "

   'RESPONSE.WRITE " ***   TRANSACAO_OBJETO ****   "
   'RESPONSE.WRITE str_SQL

   db.execute(str_SQL)

   SUB GRAVA_LOG(COD,TABELA,OPE,CONECTA)

    TX_CHAVE_TABELA=COD
    TX_TABELA=TABELA
    TX_OPERACAO=OPE

    SSQL_LOG=""
    SSQL_LOG="INSERT INTO " & Session("PREFIXO") & "LOG_GERAL(LOGE_TX_CHAVE_TABELA,LOGE_TX_TABELA,LOGE_DT_DATA_LOG,LOGE_TX_OPERACAO,LOGE_TX_USUARIO) "
    SSQL_LOG=SSQL_LOG+"VALUES('" & TX_CHAVE_TABELA & "', "
    SSQL_LOG=SSQL_LOG+"'" & TX_TABELA & "', "
    SSQL_LOG=SSQL_LOG+"GETDATE(), "
    SSQL_LOG=SSQL_LOG+"'" & TX_OPERACAO & "', "
    SSQL_LOG=SSQL_LOG+"'" & Session("CdUsuario") & "')"

    IF CONECTA=1 THEN
       DB.EXECUTE(SSQL_LOG)
    ELSE
	   CONN_DB.EXECUTE(SSQL_LOG)
    END IF

   END SUB

   'call grava_log(str_FuncPrinc,"" & Session("PREFIXO") & "MACRO_PERFIL","I",1)

   db.Close
   set db = Nothing

   response.redirect "rel_funcao_transacao.asp?selMegaProcesso=" & str_MegaProcesso & "&selFuncao=" & str_FuncPrinc & "&txtMacroPerfil=" & sequencia & "&txtNomeTecnico=" & str_PrefixoNomeTecnico & str_NomeTecnico
else
   db.Close
   set db = Nothing
   response.redirect "msg_ja_existe.asp?opt=1"
end if
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
