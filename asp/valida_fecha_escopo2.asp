<<<<<<< HEAD
<%
Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"

Server.Scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

ATUAL=request("seq")

'Atualiza Fechamento de Escopo para Decomposição(Relação Final)
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_DECOMP(FEES_CD_FECHAMENTO,TRAN_CD_TRANSACAO,ATCA_CD_ATIVIDADE_CARGA,MODU_CD_MODULO,MEPR_CD_MEGA_PROCESSO,PROC_CD_PROCESSO,ATUA_TX_OPERACAO,SUPR_CD_SUB_PROCESSO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO,RELA_NR_SEQUENCIA) "
SSQL=SSQL+"(SELECT " & ATUAL & ",TRAN_CD_TRANSACAO,ATCA_CD_ATIVIDADE_CARGA,MODU_CD_MODULO,MEPR_CD_MEGA_PROCESSO,PROC_CD_PROCESSO,'I',SUPR_CD_SUB_PROCESSO,'" & session("CdUsuario") & "', GETDATE(), RELA_NR_SEQUENCIA FROM " & Session("PREFIXO") & "RELACAO_FINAL)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=4 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

'Atualiza Fechamento de Escopo para Escopo(TABELA MODU_ATIV_TRANS_CARGA)
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_ESCOPO(FEES_CD_FECHAMENTO,TRAN_CD_TRANSACAO,ATCA_CD_ATIVIDADE_CARGA,MODU_CD_MODULO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO) "
SSQL=SSQL+"(SELECT " & ATUAL & ",TRAN_CD_TRANSACAO,ATCA_CD_ATIVIDADE_CARGA,MODU_CD_MODULO,'I','" & session("cdusuario") & "',GETDATE() FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=5 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

response.redirect "final_fecha_escopo.asp"
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
      <td width="100%">&nbsp;
        <p>&nbsp;</p>
        &nbsp;</td>
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

ATUAL=request("seq")

'Atualiza Fechamento de Escopo para Decomposição(Relação Final)
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_DECOMP(FEES_CD_FECHAMENTO,TRAN_CD_TRANSACAO,ATCA_CD_ATIVIDADE_CARGA,MODU_CD_MODULO,MEPR_CD_MEGA_PROCESSO,PROC_CD_PROCESSO,ATUA_TX_OPERACAO,SUPR_CD_SUB_PROCESSO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO,RELA_NR_SEQUENCIA) "
SSQL=SSQL+"(SELECT " & ATUAL & ",TRAN_CD_TRANSACAO,ATCA_CD_ATIVIDADE_CARGA,MODU_CD_MODULO,MEPR_CD_MEGA_PROCESSO,PROC_CD_PROCESSO,'I',SUPR_CD_SUB_PROCESSO,'" & session("CdUsuario") & "', GETDATE(), RELA_NR_SEQUENCIA FROM " & Session("PREFIXO") & "RELACAO_FINAL)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=4 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

'Atualiza Fechamento de Escopo para Escopo(TABELA MODU_ATIV_TRANS_CARGA)
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "FECHA_HISTORICO_ESCOPO(FEES_CD_FECHAMENTO,TRAN_CD_TRANSACAO,ATCA_CD_ATIVIDADE_CARGA,MODU_CD_MODULO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO) "
SSQL=SSQL+"(SELECT " & ATUAL & ",TRAN_CD_TRANSACAO,ATCA_CD_ATIVIDADE_CARGA,MODU_CD_MODULO,'I','" & session("cdusuario") & "',GETDATE() FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA)"

DB.EXECUTE(SSQL)
DB.EXECUTE("UPDATE " & Session("PREFIXO") & "FECHA_ESCOPO SET FEES_TX_CONTROLA_FECHAMENTO=5 WHERE FEES_CD_FECHAMENTO=" & ATUAL)

response.redirect "final_fecha_escopo.asp"
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
      <td width="100%">&nbsp;
        <p>&nbsp;</p>
        &nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
