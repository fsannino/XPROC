<%
micro=request("selMicro")
str_DescMicroPerfil = request("txtDescM")
str_DescMicroPerfil_Original = request("txtDescMicroPerfil_Original")
str_DescDetalhada = request("txtdetalM")
str_DescDetalhada_Original = request("txtDescDetalhada_Original")
str_Especificacao = Trim(request("txtespecM"))
str_Especificacao_Original = request("txtEspecificacao_Original")

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")
STR_ = ""
ESPEC = ucase(request("Espec_ant")) + " ***** " & ucase(request("txtespecM")) & " - " & Session("CdUsuario") & " - " & DATE
If Len(str_Especificacao) <> 0 OR Trim(str_DescMicroPerfil) <> Trim(ucase(str_DescMicroPerfil_Original)) OR Trim(str_DescDetalhada) <> Trim(ucase(str_DescDetalhada_Original)) then 
   IF Len(str_Especificacao) <> 0 THEN
      STR_ = STR_ & " Especificação -  "
   END IF
   IF Trim(str_DescMicroPerfil) <> Trim(ucase(str_DescMicroPerfil_Original)) THEN
      STR_ = STR_ & " Descrição -  "
   END IF
   IF Trim(str_DescDetalhada) <> Trim(ucase(str_DescDetalhada_Original)) THEN
      STR_ = STR_ & " Descrição detalhada -  "
   END IF
   'RESPONSE.Write(STR_)
   ssql=""
   ssql="UPDATE MICRO_PERFIL"
   ssql=ssql + " SET MICR_TX_DESC_MICRO_PERFIL='" & ucase(Trim(request("txtDescM"))) &"',"
   ssql=ssql + "MICR_TX_DESC_DETA_MICRO_PERFIL='" & ucase(Trim(request("txtdetalM"))) & "',"
   ssql=ssql + "MICR_TX_ESPECIFICACAO='" & ESPEC & "',"
   ssql=ssql + "ATUA_CD_NR_USUARIO='" & Session("CdUsuario") & "',"
   ssql=ssql + "ATUA_DT_ATUALIZACAO=GETDATE()"
   ssql=ssql+" WHERE MICR_TX_SEQ_MICRO_PERFIL='" & micro & "'"
   'RESPONSE.Write("<P> 1 " & ssql)
   conn_db.execute(ssql)
   str_SQL = ""
   str_SQL = str_SQL & " SELECT distinct MHVA_TX_SITUACAO_MICRO "
   str_SQL = str_SQL & " FROM dbo.MICRO_HISTORICO_VALIDACAO "
   str_SQL = str_SQL & " WHERE (MICR_TX_SEQ_MICRO_PERFIL = '" & micro &  "')" 
   str_SQL = str_SQL & " AND MHVA_TX_SITUACAO_MICRO IN ('CR') "
   'RESPONSE.Write("<P> 2 " & str_SQl)   
   set rdsMicroCriado = conn_db.execute(str_SQL)
   set rs=conn_db.execute("SELECT * FROM MICRO_PERFIL WHERE MICR_TX_SEQ_MICRO_PERFIL='" & micro & "'")
   if not rdsMicroCriado.EOF then
      str_Proximo_Status = "AR"
      conn_db.execute("UPDATE MICRO_PERFIL SET MICR_TX_SITUACAO='" & str_Proximo_Status & "' WHERE MICR_TX_SEQ_MICRO_PERFIL='" & micro & "'")
   else
      str_Proximo_Status = rs("MICR_TX_SITUACAO")
   end if
        		
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
   str_Comentario = " ALTERADO " & STR_		
   SSQL=""
   SSQL="INSERT INTO " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO(MHVA_NR_SEQUENCIA_HIST, MICR_TX_SEQ_MICRO_PERFIL, MHVA_TX_SITUACAO_MICRO, MHVA_TX_COMENTARIO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
   SSQL=SSQL+"VALUES(" & ATUAL &",'" & rs("MICR_TX_SEQ_MICRO_PERFIL") & "', '" & str_Proximo_Status & "','" & str_Comentario  & "','I', '" & Session("CdUsuario") & "', GETDATE())"        		
   'RESPONSE.Write("<P> 3 " & ssql)   
   CONN_DB.EXECUTE(SSQL)
else
   'RESPONSE.Write(" não faz nada ")   
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
        <p align="center"><font color="#330099" face="Verdana" size="3">ALTERAÇÃO
        DE MICRO PERFIL</font></p>
        <p align="center"><b><font face="Verdana" size="2" color="#330099">Micro
        Perfil alterado com Sucesso</font></b></p>
        
  <table width="512" border="0" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td height="41" width="176" align="right"><a href="seleciona_micro_perfil.asp?pOPT=2"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
      <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
        para tela de Alteração de Micro-Perfil</font></td>
    </tr>
    <% if str_Proximo_Status <> "AR" AND str_Proximo_Status <> "EC" AND str_Proximo_Status <> "" then %>
    <tr> 
      <td height="41" align="right"><a href="valida_alterar_micro_manda_g3.asp?pOpt=1&selMicro=<%=micro%>"><%'=str_Proximo_Status%><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
      <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Encaminha 
        para cria&ccedil;&atilde;o no R3</font></td>
    </tr>
    <% else %>
    <% end if %>
    <tr> 
      <td height="41" width="176" align="right"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
      <td height="41" width="332"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
        para tela Principal</font></td>
    </tr>
  </table>
  </form>
</body>
</html>
