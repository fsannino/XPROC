<<<<<<< HEAD
<%
Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"

Server.Scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

comentario=ucase(request("txtcoment"))

set rs=db.execute("SELECT MAX(FEES_CD_FECHAMENTO)AS CODIGO FROM " & Session("PREFIXO") & "FECHA_ESCOPO")

ATUAL=RS("CODIGO")+1

if atual>1 then
	atual=atual
else
	atual=1
end if

IF ISNULL(RS("CODIGO")) THEN
	CODIGO=0
ELSE
	CODIGO=RS("CODIGO")
END IF

SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FECHA_ESCOPO WHERE FEES_CD_FECHAMENTO= " & CODIGO)

IF TEMP.EOF=FALSE THEN
	DATA_ANT="'" & TEMP("FEES_DT_FECHAMENTO") & "'"
ELSE
	DATA_ANT="GETDATE()"
END IF

DB.EXECUTE("INSERT INTO " & Session("PREFIXO") & "FECHA_ESCOPO(FEES_CD_FECHAMENTO,FEES_DT_FECHAMENTO,FEES_DT_FECHAMENTO_ANTERIOR,FEES_TX_CHAVE_QUEM_FECHOU,FEES_TX_COMENTARIO)VALUES(" & ATUAL & ", GETDATE(), " & DATA_ANT & ",'" & session("CdUsuario") & "', '" & comentario & "' )")

'Atualiza Fechamento de Escopo para Transações
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_TRANSACAO(TRAN_TX_DESC_TRANSACAO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,TRAN_CD_TRANSACAO,ATUA_DT_ATUALIZACAO,FEES_CD_FECHAMENTO) "
SSQL=SSQL+"(SELECT TRAN_TX_DESC_TRANSACAO, 'I' , '" & Session("CdUsuario ")& "', TRAN_CD_TRANSACAO, GETDATE(), " & ATUAL & " FROM " & Session("PREFIXO") & "TRANSACAO)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=1 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

'Atualiza Fechamento de Escopo para Cenários
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_CENARIO(CENA_TX_SITUACAO_VALIDACAO, CENA_DT_VALIDACAO, CENA_TX_TITULO_CENARIO,CENA_TX_DESC_CENARIO,PROC_CD_PROCESSO,SUPR_CD_SUB_PROCESSO,ONDA_CD_ONDA,CLCE_CD_NR_CLASSE_CENARIO,MEPR_CD_MEGA_PROCESSO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO,CENA_TX_SITUACAO,CENA_NR_SEQUENCIA,CENA_NR_SEQUENCIA_ORDEM,CENA_TX_SITUACAO_LOTUS,CENA_TX_SAIDA,CENA_TX_ENTRADA,CENA_CD_CENARIO,CENA_TX_GERA_FLUXO,CENA_TX_SITUACAO_TESTE,CENA_TX_SITU_DESENHO_TIPO,CENA_TX_SITU_DESENHO_CONF,CENA_TX_SITU_DESENHO_DESE,CENA_TX_CD_CENARIO,CENA_TX_EMPRESA_RELAC,CENA_TX_SITU_DESENHO_TESTE,FEES_CD_FECHAMENTO) "
SSQL=SSQL+"(SELECT CENA_TX_SITUACAO_VALIDACAO, CENA_DT_VALIDACAO, CENA_TX_TITULO_CENARIO,CENA_TX_DESC_CENARIO,PROC_CD_PROCESSO,SUPR_CD_SUB_PROCESSO,ONDA_CD_ONDA,CLCE_CD_NR_CLASSE_CENARIO,MEPR_CD_MEGA_PROCESSO,'I','" & Session("CdUsuario") & "',GETDATE(),CENA_TX_SITUACAO,CENA_NR_SEQUENCIA,CENA_NR_SEQUENCIA_ORDEM,CENA_TX_SITUACAO_LOTUS,CENA_TX_SAIDA,CENA_TX_ENTRADA,CENA_CD_CENARIO,CENA_TX_GERA_FLUXO,CENA_TX_SITUACAO_TESTE,CENA_TX_SITU_DESENHO_TIPO,CENA_TX_SITU_DESENHO_CONF,CENA_TX_SITU_DESENHO_DESE,CENA_TX_CD_CENARIO,CENA_TX_EMPRESA_RELAC,CENA_TX_SITU_DESENHO_TESTE," & ATUAL & " FROM " & Session("PREFIXO") & "CENARIO)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=2 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

'Atualiza Fechamento de Escopo para Cenários x Transações
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_CENA_TRANS(FEES_CD_FECHAMENTO,CENA_CD_CENARIO,CETR_NR_SEQUENCIA,TRAN_CD_TRANSACAO,OPES_CD_OPERACAO_ESP,BPPP_CD_BPP,MEPR_CD_MEGA_PROCESSO,CENA_CD_CENARIO_SEGUINTE,CENA_NR_SEQUENCIA_TRANS,CETR_TX_DESC_TRANSACAO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO,CETR_TX_TIPO_RELACAO,DESE_CD_DESENVOLVIMENTO,CETR_TX_SITU_TESTE_CENARIO,CETR_TX_RESP_TESTE_CENARIO) "
SSQL=SSQL+"(SELECT " & ATUAL & ",CENA_CD_CENARIO,CETR_NR_SEQUENCIA,TRAN_CD_TRANSACAO,OPES_CD_OPERACAO_ESP,BPPP_CD_BPP,MEPR_CD_MEGA_PROCESSO,CENA_CD_CENARIO_SEGUINTE,CENA_NR_SEQUENCIA_TRANS,CETR_TX_DESC_TRANSACAO,'I','" & session("cdusuario") & "',GETDATE(),CETR_TX_TIPO_RELACAO,DESE_CD_DESENVOLVIMENTO,CETR_TX_SITU_TESTE_CENARIO,CETR_TX_RESP_TESTE_CENARIO FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=3 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

response.redirect "valida_fecha_escopo2.asp?SEQ=" & ATUAL

%>
<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="valida_fecha_escopo.asp">
      <table width="812" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="140" height="20" colspan="2">&nbsp;</td>
      <td width="1065" height="60" colspan="2">&nbsp;</td>
      <td width="1" valign="top"> 
        <table width="153" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="43" valign="middle" align="right"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="25" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="34" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="43" valign="middle" align="right"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="25" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="34" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="2">&nbsp; </td>
      <td height="20" width="136">&nbsp;</td>
      <td height="20" width="27"> 
        <p align="center">&nbsp;
      </td>
      <td height="20" width="264" colspan="2">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="100%">&nbsp;</td>
    </tr>
  </table> 
</form>
</body>
</html>
=======
<%
Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"

Server.Scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

comentario=ucase(request("txtcoment"))

set rs=db.execute("SELECT MAX(FEES_CD_FECHAMENTO)AS CODIGO FROM " & Session("PREFIXO") & "FECHA_ESCOPO")

ATUAL=RS("CODIGO")+1

if atual>1 then
	atual=atual
else
	atual=1
end if

IF ISNULL(RS("CODIGO")) THEN
	CODIGO=0
ELSE
	CODIGO=RS("CODIGO")
END IF

SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FECHA_ESCOPO WHERE FEES_CD_FECHAMENTO= " & CODIGO)

IF TEMP.EOF=FALSE THEN
	DATA_ANT="'" & TEMP("FEES_DT_FECHAMENTO") & "'"
ELSE
	DATA_ANT="GETDATE()"
END IF

DB.EXECUTE("INSERT INTO " & Session("PREFIXO") & "FECHA_ESCOPO(FEES_CD_FECHAMENTO,FEES_DT_FECHAMENTO,FEES_DT_FECHAMENTO_ANTERIOR,FEES_TX_CHAVE_QUEM_FECHOU,FEES_TX_COMENTARIO)VALUES(" & ATUAL & ", GETDATE(), " & DATA_ANT & ",'" & session("CdUsuario") & "', '" & comentario & "' )")

'Atualiza Fechamento de Escopo para Transações
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_TRANSACAO(TRAN_TX_DESC_TRANSACAO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,TRAN_CD_TRANSACAO,ATUA_DT_ATUALIZACAO,FEES_CD_FECHAMENTO) "
SSQL=SSQL+"(SELECT TRAN_TX_DESC_TRANSACAO, 'I' , '" & Session("CdUsuario ")& "', TRAN_CD_TRANSACAO, GETDATE(), " & ATUAL & " FROM " & Session("PREFIXO") & "TRANSACAO)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=1 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

'Atualiza Fechamento de Escopo para Cenários
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_CENARIO(CENA_TX_SITUACAO_VALIDACAO, CENA_DT_VALIDACAO, CENA_TX_TITULO_CENARIO,CENA_TX_DESC_CENARIO,PROC_CD_PROCESSO,SUPR_CD_SUB_PROCESSO,ONDA_CD_ONDA,CLCE_CD_NR_CLASSE_CENARIO,MEPR_CD_MEGA_PROCESSO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO,CENA_TX_SITUACAO,CENA_NR_SEQUENCIA,CENA_NR_SEQUENCIA_ORDEM,CENA_TX_SITUACAO_LOTUS,CENA_TX_SAIDA,CENA_TX_ENTRADA,CENA_CD_CENARIO,CENA_TX_GERA_FLUXO,CENA_TX_SITUACAO_TESTE,CENA_TX_SITU_DESENHO_TIPO,CENA_TX_SITU_DESENHO_CONF,CENA_TX_SITU_DESENHO_DESE,CENA_TX_CD_CENARIO,CENA_TX_EMPRESA_RELAC,CENA_TX_SITU_DESENHO_TESTE,FEES_CD_FECHAMENTO) "
SSQL=SSQL+"(SELECT CENA_TX_SITUACAO_VALIDACAO, CENA_DT_VALIDACAO, CENA_TX_TITULO_CENARIO,CENA_TX_DESC_CENARIO,PROC_CD_PROCESSO,SUPR_CD_SUB_PROCESSO,ONDA_CD_ONDA,CLCE_CD_NR_CLASSE_CENARIO,MEPR_CD_MEGA_PROCESSO,'I','" & Session("CdUsuario") & "',GETDATE(),CENA_TX_SITUACAO,CENA_NR_SEQUENCIA,CENA_NR_SEQUENCIA_ORDEM,CENA_TX_SITUACAO_LOTUS,CENA_TX_SAIDA,CENA_TX_ENTRADA,CENA_CD_CENARIO,CENA_TX_GERA_FLUXO,CENA_TX_SITUACAO_TESTE,CENA_TX_SITU_DESENHO_TIPO,CENA_TX_SITU_DESENHO_CONF,CENA_TX_SITU_DESENHO_DESE,CENA_TX_CD_CENARIO,CENA_TX_EMPRESA_RELAC,CENA_TX_SITU_DESENHO_TESTE," & ATUAL & " FROM " & Session("PREFIXO") & "CENARIO)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=2 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

'Atualiza Fechamento de Escopo para Cenários x Transações
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_CENA_TRANS(FEES_CD_FECHAMENTO,CENA_CD_CENARIO,CETR_NR_SEQUENCIA,TRAN_CD_TRANSACAO,OPES_CD_OPERACAO_ESP,BPPP_CD_BPP,MEPR_CD_MEGA_PROCESSO,CENA_CD_CENARIO_SEGUINTE,CENA_NR_SEQUENCIA_TRANS,CETR_TX_DESC_TRANSACAO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO,CETR_TX_TIPO_RELACAO,DESE_CD_DESENVOLVIMENTO,CETR_TX_SITU_TESTE_CENARIO,CETR_TX_RESP_TESTE_CENARIO) "
SSQL=SSQL+"(SELECT " & ATUAL & ",CENA_CD_CENARIO,CETR_NR_SEQUENCIA,TRAN_CD_TRANSACAO,OPES_CD_OPERACAO_ESP,BPPP_CD_BPP,MEPR_CD_MEGA_PROCESSO,CENA_CD_CENARIO_SEGUINTE,CENA_NR_SEQUENCIA_TRANS,CETR_TX_DESC_TRANSACAO,'I','" & session("cdusuario") & "',GETDATE(),CETR_TX_TIPO_RELACAO,DESE_CD_DESENVOLVIMENTO,CETR_TX_SITU_TESTE_CENARIO,CETR_TX_RESP_TESTE_CENARIO FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=3 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

response.redirect "valida_fecha_escopo2.asp?SEQ=" & ATUAL

%>
<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="valida_fecha_escopo.asp">
      <table width="812" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="140" height="20" colspan="2">&nbsp;</td>
      <td width="1065" height="60" colspan="2">&nbsp;</td>
      <td width="1" valign="top"> 
        <table width="153" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="43" valign="middle" align="right"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="25" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="34" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="43" valign="middle" align="right"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="25" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="34" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="2">&nbsp; </td>
      <td height="20" width="136">&nbsp;</td>
      <td height="20" width="27"> 
        <p align="center">&nbsp;
      </td>
      <td height="20" width="264" colspan="2">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="100%">&nbsp;</td>
    </tr>
  </table> 
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
