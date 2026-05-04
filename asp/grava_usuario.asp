<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%>
<%
str_Cod = UCase(Request("txtCodUsuario"))


str_NomeUsuario = Request("txtNomeUsuario")
str_Categoria = Request("rdbCategoria")
str_CdUsuario = UCase(Request("Cd"))
'response.Write(str_CdUsuario)
'response.End()
str_email = Request("txtemail2")
str_senha = Request("txtsenha")

if Request("txtMegaSelecionado") = "" then
   str_MegaSelecionado = "0"
else
   str_MegaSelecionado = Request("txtMegaSelecionado")
end if   

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Sql = ""
str_Sql = str_Sql & " SELECT USUA_TX_NOME_USUARIO"
str_Sql = str_Sql & " FROM   dbo.USUARIO"
str_Sql = str_Sql & " WHERE  USUA_CD_USUARIO = '" & str_CdUsuario &"'"

set rdsExisteUsu = db.execute(str_Sql)

if not rdsExisteUsu.Eof then
	str_Url = "msg_erro.asp?pCdMsgErro=1&Nome=" & rdsExisteUsu("USUA_TX_NOME_USUARIO")
	response.Redirect(str_Url)
end if



ssql=""
ssql=" INSERT INTO " & Session("PREFIXO") & "USUARIO ("
ssql=ssql & " USUA_CD_USUARIO, USUA_TX_NOME_USUARIO, "
ssql=ssql & " USUA_TX_CATEGORIA, ATUA_TX_OPERACAO, "
ssql=ssql & " ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, "
ssql=ssql & " USUA_TX_EMAIL_EXTERNO, USUA_TX_SENHA "
ssql=ssql & " ) VALUES('" & ucase(str_CdUsuario) & "','" & ucase(str_NomeUsuario) & "','" & str_Categoria & "',"
ssql=ssql & " 'C', '" & Session("CdUsuario") & "', GETDATE(), '" & lcase(str_email) & "','" & lcase(str_senha) & "')"

db.execute(ssql)

strChave = CStr(ucase(str_CdUsuario)) 

if err.number=0 then
	erro=0
else
	erro=1
end if

'*********************************************************
db.execute("DELETE FROM " & Session("PREFIXO") & "ACESSO WHERE USUA_CD_USUARIO='" & strChave & "'")

str_SQL_Novo_Acesso = ""

'guarda o conteúdo da String
str_valor = str_MegaSelecionado

response.Write(strChave)
response.Write(str_valor)

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
	   call Grava_Novo_Acesso(strChave, str_atual)
	   valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop

Sub Grava_Novo_Acesso(strU, strMP)
	str_SQL_Novo_Acesso = ""
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " INSERT INTO " & Session("PREFIXO") & "ACESSO ( "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " USUA_CD_USUARIO, MEPR_CD_MEGA_PROCESSO, "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " ATUA_DT_ATUALIZACAO) VALUES ('"
   	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & strU & "'," & strMP & ", "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
   	response.write str_SQL_Nova_Sub_Empr
	Set rdsNovo = db.Execute(str_SQL_Novo_Acesso)
	strChave = CStr(strU) & " " & CStr(strMP) '& " " & CStr(strSP) & " " & CStr(strM) & " " & CStr(strA) & CStr(strT)	
end sub

db.Close
set db = Nothing

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
<form method="POST" action="valida_altera_processo.asp?mega=<%=str_MegaProcesso%>&Proc=<%=str_Processo%>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="3%">&nbsp;</td>
      <td height="20" width="43%">
<%'=ssql%></td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="39%">
<%'=str_CdUsuario%></td>
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
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Cadastro 
        de Usu&aacute;rio</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <%if erro=0 then%>
    <tr> 
      <td width="3%"></td>
      <td width="24%"></td>
      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><b>Usu&aacute;rio 
        Cadastrado com Sucesso!</b></font></td>
      <td width="14%"></td>
    </tr>
    <%else%>
    <tr> 
      <td width="3%"></td>
      <td width="24%"></td>
      <td width="59%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#800000">Não 
        foi Possível Incluir o Registro. Usu&aacute;rio j&aacute; cadastrado.</font></b></td>
      <td width="14%"></td>
    </tr>
    <%end if%>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">
        <table width="74%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr> 
            <td height="41"><a href="cad_usuario.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela de Cadastro de Usu&aacute;rio</font></td>
          </tr>
          <tr> 
            <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
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
<p>&nbsp;</p><!-- #EndEditable -->
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%>
<%
str_Cod = UCase(Request("txtCodUsuario"))


str_NomeUsuario = Request("txtNomeUsuario")
str_Categoria = Request("rdbCategoria")
str_CdUsuario = UCase(Request("Cd"))
'response.Write(str_CdUsuario)
'response.End()
str_email = Request("txtemail2")
str_senha = Request("txtsenha")

if Request("txtMegaSelecionado") = "" then
   str_MegaSelecionado = "0"
else
   str_MegaSelecionado = Request("txtMegaSelecionado")
end if   

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Sql = ""
str_Sql = str_Sql & " SELECT USUA_TX_NOME_USUARIO"
str_Sql = str_Sql & " FROM   dbo.USUARIO"
str_Sql = str_Sql & " WHERE  USUA_CD_USUARIO = '" & str_CdUsuario &"'"

set rdsExisteUsu = db.execute(str_Sql)

if not rdsExisteUsu.Eof then
	str_Url = "msg_erro.asp?pCdMsgErro=1&Nome=" & rdsExisteUsu("USUA_TX_NOME_USUARIO")
	response.Redirect(str_Url)
end if



ssql=""
ssql=" INSERT INTO " & Session("PREFIXO") & "USUARIO ("
ssql=ssql & " USUA_CD_USUARIO, USUA_TX_NOME_USUARIO, "
ssql=ssql & " USUA_TX_CATEGORIA, ATUA_TX_OPERACAO, "
ssql=ssql & " ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, "
ssql=ssql & " USUA_TX_EMAIL_EXTERNO, USUA_TX_SENHA "
ssql=ssql & " ) VALUES('" & ucase(str_CdUsuario) & "','" & ucase(str_NomeUsuario) & "','" & str_Categoria & "',"
ssql=ssql & " 'C', '" & Session("CdUsuario") & "', GETDATE(), '" & lcase(str_email) & "','" & lcase(str_senha) & "')"

db.execute(ssql)

strChave = CStr(ucase(str_CdUsuario)) 

if err.number=0 then
	erro=0
else
	erro=1
end if

'*********************************************************
db.execute("DELETE FROM " & Session("PREFIXO") & "ACESSO WHERE USUA_CD_USUARIO='" & strChave & "'")

str_SQL_Novo_Acesso = ""

'guarda o conteúdo da String
str_valor = str_MegaSelecionado

response.Write(strChave)
response.Write(str_valor)

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
	   call Grava_Novo_Acesso(strChave, str_atual)
	   valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop

Sub Grava_Novo_Acesso(strU, strMP)
	str_SQL_Novo_Acesso = ""
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " INSERT INTO " & Session("PREFIXO") & "ACESSO ( "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " USUA_CD_USUARIO, MEPR_CD_MEGA_PROCESSO, "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " ATUA_DT_ATUALIZACAO) VALUES ('"
   	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & strU & "'," & strMP & ", "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
   	response.write str_SQL_Nova_Sub_Empr
	Set rdsNovo = db.Execute(str_SQL_Novo_Acesso)
	strChave = CStr(strU) & " " & CStr(strMP) '& " " & CStr(strSP) & " " & CStr(strM) & " " & CStr(strA) & CStr(strT)	
end sub

db.Close
set db = Nothing

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
<form method="POST" action="valida_altera_processo.asp?mega=<%=str_MegaProcesso%>&Proc=<%=str_Processo%>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="3%">&nbsp;</td>
      <td height="20" width="43%">
<%'=ssql%></td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="39%">
<%'=str_CdUsuario%></td>
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
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Cadastro 
        de Usu&aacute;rio</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <%if erro=0 then%>
    <tr> 
      <td width="3%"></td>
      <td width="24%"></td>
      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><b>Usu&aacute;rio 
        Cadastrado com Sucesso!</b></font></td>
      <td width="14%"></td>
    </tr>
    <%else%>
    <tr> 
      <td width="3%"></td>
      <td width="24%"></td>
      <td width="59%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#800000">Não 
        foi Possível Incluir o Registro. Usu&aacute;rio j&aacute; cadastrado.</font></b></td>
      <td width="14%"></td>
    </tr>
    <%end if%>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">
        <table width="74%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr> 
            <td height="41"><a href="cad_usuario.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela de Cadastro de Usu&aacute;rio</font></td>
          </tr>
          <tr> 
            <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
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
<p>&nbsp;</p><!-- #EndEditable -->
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
