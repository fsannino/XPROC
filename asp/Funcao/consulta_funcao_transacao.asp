 

<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO"
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO INNER JOIN " & Session("PREFIXO") & "TRANSACAO ON " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
'str_SQL_Transacao = str_SQL_Transacao & " WHERE FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = 1"
'str_SQL_Transacao = str_SQL_Transacao & " AND FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = 1  "
'str_SQL_Transacao = str_SQL_Transacao & " AND FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = 1  "
'str_SQL_Transacao = str_SQL_Transacao & " AND FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = 1 " 
'str_SQL_Transacao = str_SQL_Transacao & " AND FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '1'"
str_SQL_Transacao = str_SQL_Transacao & " order by " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO "
Set rdsTransacao = Conn_db.Execute(str_SQL_Transacao)

str_SQL_Funcao = ""
str_SQL_Funcao = str_SQL_Funcao & " SELECT " 
str_SQL_Funcao = str_SQL_Funcao & " TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO, "
str_SQL_Funcao = str_SQL_Funcao & " PROC_CD_PROCESSO, SUPR_CD_SUB_PROCESSO, "
str_SQL_Funcao = str_SQL_Funcao & " ATCA_CD_ATIVIDADE_CARGA, "
str_SQL_Funcao = str_SQL_Funcao & " FUNE_CD_FUNCAO_NEGOCIO"
str_SQL_Funcao = str_SQL_Funcao & " FROM FUN_NEG_TRANSACAO"
'str_SQL_Funcao = str_SQL_Funcao & " WHERE FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = 1"
'str_SQL_Funcao = str_SQL_Funcao & " AND FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = 1  "
'str_SQL_Funcao = str_SQL_Funcao & " AND FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = 1  "
'str_SQL_Funcao = v & " AND FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = 1 " 
'str_SQL_Funcao = str_SQL_Funcao & " AND FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '1'"
str_SQL_Funcao = str_SQL_Funcao & " order by " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO "
'Set rdsFuncao = Conn_db.Execute(str_SQL_Funcao)
 
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
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
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
          <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="795" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="72">&nbsp;</td>
    <td width="647">&nbsp;</td>
    <td width="76">&nbsp;</td>
  </tr>
  <tr>
    <td width="72">&nbsp;</td>
    <td width="647">&nbsp;</td>
    <td width="76">&nbsp;</td>
  </tr>
  <tr>
    <td width="72">&nbsp;</td>
    <td width="647">&nbsp;</td>
    <td width="76">&nbsp;</td>
  </tr>
</table>
<table width="45%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="30%">&nbsp;</td>
    <td width="37%">&nbsp;</td>
    <td width="33%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="30%">&nbsp;</td>
    <td width="37%">&nbsp;</td>
    <td width="33%">&nbsp;</td>
  </tr>
  <tr bgcolor="#0000FF"> 
    <td width="30%"> 
      <div align="center"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o</font></b></div>
    </td>
    <td width="37%"> 
      <div align="center"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></b></div>
    </td>
    <td width="33%"> 
      <div align="center"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Funcao</font></b></div>
    </td>
  </tr>
  <tr> 
    <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Ativ 
      &gt; <%=rdsTransacao("TRAN_CD_TRANSACAO")%></font></td>
    <td width="33%">&nbsp;</td>
  </tr>
  <tr> 
    <% do While not rdsTransacao.EOF  %>
    <td width="30%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsTransacao("TRAN_CD_TRANSACAO")%> </font></td>
    <td width="37%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsTransacao("TRAN_TX_DESC_TRANSACAO")%></font></td>
    <td width="33%">&nbsp;</td>
  </tr>
  <% rdsTransacao.movenext
  Loop 
  rdsTransacao.Close
  set rdsTransacao = Nothing
  %>
  <tr> 
    <td width="30%">&nbsp;</td>
    <td width="37%">&nbsp;</td>
    <td width="33%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
