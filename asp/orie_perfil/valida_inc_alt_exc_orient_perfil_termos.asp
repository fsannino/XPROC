<%

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

dim int_Tot_Orientacoes 
int_Tot_Orientacoes = 0
if request("txtOpt") <> "" then
   str_Opt = request("txtOpt")
else
   str_Opt = "0"
end if   
'response.Write(str_Opt)
if str_Opt = "I" then
   dim str_Orientacoes(4)
   dim str_Termos(4)
   str_Orientacoes(1) = request("txtOrientacoes1")
   str_Orientacoes(2) = request("txtOrientacoes2")
   str_Orientacoes(3) = request("txtOrientacoes3")
   str_Orientacoes(4) = request("txtOrientacoes4")
   str_Termos(1) = request("txtTermo1")
   str_Termos(2) = request("txtTermo2")
   str_Termos(3) = request("txtTermo3")
   str_Termos(4) = request("txtTermo4")
   atual=0
   set rs=db.execute("SELECT MAX(ORTE_NR_SEQUENCIAL) AS CODIGO FROM  PERFIL_ORIEN_GERAL_TERMOS")
   if not isnull(rs("CODIGO")) then
	  atual = rs("CODIGO")
   end if
   if atual=0 then
	  atual=0
   else
	  atual=atual
   end if
   For i = 1 to 4
       if str_Orientacoes(i) <> "" and str_Termos(i) <> "" then
	      int_Tot_Orientacoes = int_Tot_Orientacoes + 1
		  atual = atual + 1
          ssql=""
          ssql="INSERT INTO " & Session("PREFIXO") & "PERFIL_ORIEN_GERAL_TERMOS ("
          ssql=ssql & " ORTE_NR_SEQUENCIAL, "
          ssql=ssql & " ORTE_TX_TERMO, "
          ssql=ssql & " ORTE_TX_DESCRICAO, ATUA_TX_OPERACAO, "
          ssql=ssql & " ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO "
          ssql=ssql & ") VALUES (" & atual & ", "
          ssql=ssql+"'" & str_Termos(i) & "',"
          ssql=ssql+"'" & str_Orientacoes(i) & "',"
          ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE())"
          'response.write ssql
          db.execute(ssql)
       end if
   next   
elseif str_Opt = "A" then
   if request("txtCdTermo") <> "" then
      str_CdTermo = request("txtCdTermo")
   else
      str_CdTermo = "0"
   end if  
   str_Termo = request("txtTermo1") 
   str_Orientacoes = request("txtOrientacoes1") 
   if str_Orientacoes <> "" then
      int_Tot_Orientacoes = int_Tot_Orientacoes + 1
      ssql=""
      ssql=ssql & "UPDATE PERFIL_ORIEN_GERAL_TERMOS SET "
      ssql=ssql & " ORTE_TX_TERMO = '" & str_Termo & "' "	  	  
      ssql=ssql & " ,ORTE_TX_DESCRICAO = '" & str_Orientacoes & "' "
      ssql=ssql & " ,ATUA_TX_OPERACAO = 'A' "
      ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "' "
      ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
      ssql=ssql & " WHERE ORTE_NR_SEQUENCIAL = " & str_CdTermo
	  'response.Write(ssql)
      db.execute(ssql)	
   end if  
elseif str_Opt = "E" then
   if request("txtCdTermo") <> "" then
      str_CdTermo = request("txtCdTermo")
   else
      str_CdTermo = "0"
   end if  
   if str_CdTermo <> "" then
      int_Tot_Orientacoes = int_Tot_Orientacoes + 1
      ssql=""
      ssql=ssql & "DELETE FROM PERFIL_ORIEN_GERAL_TERMOS "
      ssql=ssql & " WHERE ORTE_NR_SEQUENCIAL = " & str_CdTermo
	  'response.Write(ssql)
      db.execute(ssql)	
   end if  
end if

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body topmargin="0" leftmargin="0">
<form method="POST" action="" name="frm1">
<input type="hidden" name="txtpub" size="20"><input type="hidden" name="txtQua" size="20">
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
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"></td>
          <td width="26"></td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%
if int_Tot_Orientacoes > 1 then
   str_Texto = " termos"
else
   str_Texto = " termo"
end if
%>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        
  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Cadastro 
    de Termos</font></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2"> 
    Registro incluído com sucesso: Total de </font><font face="Verdana" color="#330099" size="3"> 
    <%=int_Tot_Orientacoes%> <%=str_Texto%></font></b></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
  <table border="0" width="889" height="162">
    <tr> 
      <td width="287" height="37"></td>
      <td width="26" height="37"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
      <td height="37" width="556"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
        para tela Principal</font></td>
    </tr>
	<% if str_Opt = "I" then %>
    <tr> 
      <td width="287" height="37"></td>
      <td width="26" height="37"> <p align="right"><a href="inclui_ori_gerais_mape_perfil_termos.asp"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
      <td height="37" width="556"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar 
        para a tela de Inclus&atilde;o de Termos</font></td>
    </tr>
	<% elseif str_Opt = "A" then %>
    <tr> 
      <td width="287" height="27"></td>
      <td height="37"> <p align="right"><a href="seleciona_ori_gerais_mape_perfil_termos.asp?pOpt=A"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
      <td height="37"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar 
        para a tela de Altera&ccedil;&atilde;o de Termos</font></td>
    </tr>
	<% elseif str_Opt = "E" then %>
    <tr> 
      <td height="27"></td>
      <td height="37"> <p align="right"><a href="seleciona_ori_gerais_mape_perfil_termos.asp?pOpt=E"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
      <td height="37"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar 
        para a tela de Exclus&atilde;o de Termos</font></td>
    </tr>
	<% end if %>
  </table>
  </form>

<p>&nbsp;</p>

</body>

</html>

