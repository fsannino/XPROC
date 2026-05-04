<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Acao = request("txtAcao")
str_Cd_Evento = request("SelEvento")

If str_Acao <> "E" then
   str_DiaDtEvento=Right("00" & request("selDiaDtEvento"),2)
   str_MesDtEvento=Right("00" & request("selMesDtEvento"),2)
   str_AnoDtEvento=Right("00" & request("selAnoDtEvento"),2)

   str_Data = str_DiaDtEvento & "/" & str_MesDtEvento & "/" & str_AnoDtEvento
   str_Data_DB = str_MesDtEvento  & "/" & str_DiaDtEvento & "/" & str_AnoDtEvento 
   'response.Write(str_Data)
   str_Desc = UCase(request("txtDesc"))
   if str_MesDtEvento = 2 then
      if str_DiaDtEvento > 28 then
         response.redirect "envia_msg_tela.asp?opt=0&acao=" & str_Acao
      end if
   end if

   if InStrRev("04/06/09/11", str_MesDtEvento) <> 0 then
      if str_DiaDtEvento > 30 then
         response.redirect "envia_msg_tela.asp?opt=1&acao=" & str_Acao
      end if
   end if
   if IsDate(CDate(str_Data)) = false then
      response.redirect "envia_msg_tela.asp?opt=2&acao=" & str_Acao
   end if
 
   If str_Acao = "I" then
      set rs=db.execute("SELECT MAX(EVEN_NR_SEQUENCIAL) AS CAMPO FROM " & Session("PREFIXO") & "EVENTO")
      IF IsNull(rs("campo")) then
         ls_int_Proximo = 1
      else
        ls_int_Proximo =  rs("campo") + 1  
      end if

      str_SQl = ""
      str_SQL = str_SQL & " INSERT INTO EVENTO("
      str_SQL = str_SQL & " EVEN_NR_SEQUENCIAL"
      str_SQL = str_SQL & " , EVEN_DT_EVENTO "
      str_SQL = str_SQL & " , EVEN_TX_DESCRICAO"
      str_SQL = str_SQL & " , ATUA_TX_OPERACAO "
      str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
      str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
      str_SQL = str_SQL & " ) VALUES ( "
      str_SQL = str_SQL & ls_int_Proximo & " ,'" &  str_Data_DB & "','" & str_Desc & "',"
      str_SQL = str_SQL & "'C', '"& Session("CdUsuario") &"', GETDATE())"
   else
      str_SQl = ""
      str_SQL = str_SQL & " UPDATE EVENTO SET"
	  str_SQL = str_SQL & " EVEN_DT_EVENTO ='" & str_Data_DB & "'"
	  str_SQL = str_SQL & " ,EVEN_TX_DESCRICAO = '" & str_Desc & "'" 
      str_SQL = str_SQL & " ,ATUA_TX_OPERACAO = 'A'"   
      str_SQL = str_SQL & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
      str_SQL = str_SQL & " ,ATUA_DT_ATUALIZACAO = GETDATE()"
      str_SQL = str_SQL & " WHERE EVEN_NR_SEQUENCIAL=" & str_Cd_Evento	  
   end if   	  
else
   str_SQl = ""
   str_SQL = str_SQL & " DELETE FROM EVENTO "
   str_SQL = str_SQL & " WHERE EVEN_NR_SEQUENCIAL=" & str_Cd_Evento	  
end if   
'response.Write(str_SQL)
db.execute(str_SQL)
   
%>
<html>
<head>
<title>Grava cadastro de Evento</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
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
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
        
<table width="100%" border="0">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><div align="center"><font face="Verdana" color="#330099" size="3"><%="Cadastro de Evento"%></font></div></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        
<p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2">O 
  Registro foi atualizado com sucesso!</font></b></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="782">
  <tr> 
    <td width="288" height="37"></td>
    <td width="45" height="37"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
    <td height="37" width="435"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
      para tela Principal</font></td>
  </tr>
  <% if str_Acao = "I" then %>
  <tr> 
    <td width="288" height="37"></td>
    <td width="45" height="37"> <p align="right"><a href="../Escopo/incluir_evento.asp?ID=I"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
    <td height="37" width="435"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar 
      para a tela de Inclus&atilde;o de Eventos</font></td>
  </tr>
  <% end if %>
  <% if str_Acao = "A" then %>
  <tr> 
    <td width="288" height="21"></td>
    <td height="37"> <p align="right"><a href="../Escopo/seleciona_evento.asp?ID=A"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
    <td height="37"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar 
      para a tela de Altera&ccedil;&atilde;o de Eventos</font></td>
  </tr>
    <% end if %>
	<% if str_Acao = "E" then %>
  <tr> 
    <td height="21"></td>
    <td height="37"> <p align="right"><a href="../Escopo/seleciona_evento.asp?ID=E"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
    <td height="37"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar 
      para a tela de Exclus&atilde;o de Eventos</font></td>
  </tr>
  <% end if %>
</table>
<p>&nbsp;</p>
</body>

</html>
