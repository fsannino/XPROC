<%@LANGUAGE="VBSCRIPT"%> 
<%
micro=request("selMicro")

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQL = str_SQL & " SELECT distinct MHVA_TX_SITUACAO_MICRO "
str_SQL = str_SQL & " FROM dbo.MICRO_HISTORICO_VALIDACAO "
str_SQL = str_SQL & " WHERE (MICR_TX_SEQ_MICRO_PERFIL = '" & micro &  "')" 
str_SQL = str_SQL & " AND MHVA_TX_SITUACAO_MICRO IN ('CR') "
'RESPONSE.Write("<P> 2 " & str_SQl)   
set rdsMicroCriado = conn_db.execute(str_SQL)
if not rdsMicroCriado.EOF then
   str_Proximo_Status = "ER"
   conn_db.execute("UPDATE MICRO_PERFIL SET MICR_TX_SITUACAO='" & str_Proximo_Status & "' WHERE MICR_TX_SEQ_MICRO_PERFIL='" & micro & "'")
   set rs=conn_db.execute("SELECT * FROM MICRO_PERFIL WHERE MICR_TX_SEQ_MICRO_PERFIL='" & micro & "'")        		
   SSQL=""
   SSQL="SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO WHERE MICR_TX_SEQ_MICRO_PERFIL='" & rs("MICR_TX_SEQ_MICRO_PERFIL") & "'"		
   SET HIST = CONN_DB.EXECUTE(SSQL)		
   ATUAL = HIST("CODIGO")
   ATUAL = ATUAL + 1		
   if atual > 1 then
	  atual = atual
   else
	  atual=1
   end if
   str_Comentario = " EM EXCLUSÃO NO R/3 "
   SSQL=""
   SSQL="INSERT INTO " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO(MHVA_NR_SEQUENCIA_HIST, MICR_TX_SEQ_MICRO_PERFIL, MHVA_TX_SITUACAO_MICRO,MHVA_TX_COMENTARIO , ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
   SSQL=SSQL+"VALUES(" & ATUAL &",'" & rs("MICR_TX_SEQ_MICRO_PERFIL") & "', 'ER','" & str_Comentario & "','I', '" & Session("CdUsuario") & "', GETDATE())"		
   CONN_DB.EXECUTE(SSQL)   
else
   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO " 
   str_SQl = str_SQL & " Where MICR_TX_SEQ_MICRO_PERFIL = '" & micro & "'"
   CONN_DB.execute(str_SQl)    
   str_SQl = ""
   str_SQl = str_SQL & " Delete from " & Session("PREFIXO") & "MICRO_PERFIL " 
   str_SQl = str_SQL & " Where MICR_TX_SEQ_MICRO_PERFIL = '" & micro & "'"
   CONN_DB.execute(str_SQl)   
end if
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60"><font size="1" color="#FFFFFF"><b><%=Conn_String_Cogest_Gravacao%></b></font></td>
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
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://http://www.sinergia.petrobras.com.br/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
    <td colspan="3" height="20">&nbsp;
      
    </td>
  </tr>
</table>
        <p align="center"><font color="#330099" face="Verdana" size="3">EXCLUSÃO
        DE MICRO PERFIL</font></p>
        <p align="center"><b><font face="Verdana" size="2" color="#330099">Micro
        Perfil excluído com Sucesso</font></b></p>
        <table width="512" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td height="41" width="176" align="right"><a href="seleciona_micro_perfil.asp?pOPT=3"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
            <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela de Exclusão de Micro-Perfil</font></td>
          </tr>
          <tr>
            <td height="41" width="176" align="right"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
          </tr>
        </table>
  </form>
</body>
</html>
