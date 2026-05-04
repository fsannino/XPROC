<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

if request("selUsuario") = "" then
   str_Usuario=0
else
   str_Usuario=Request("selUsuario")
end if
if Request("txtMegaSelecionado") = "" then
   str_MegaSelecionado = "0"
else
   str_MegaSelecionado = Request("txtMegaSelecionado")
end if   
'response.write str_Atividade

conn_db.execute("DELETE FROM " & Session("PREFIXO") & "ACESSO WHERE USUA_CD_USUARIO='" & str_Usuario & "'")
strChave = CStr(str_Usuario) '& CStr(str_Processo) &  CStr(str_SubProcesso) & CStr(str_Empresa) 
''call grava_log(strChave,"ACESSO","D",0)	

str_SQL_Novo_Acesso = ""
Sub Grava_Novo_Acesso(strU, strMP)
	str_SQL_Novo_Acesso = ""
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " INSERT INTO " & Session("PREFIXO") & "ACESSO ( "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " USUA_CD_USUARIO, MEPR_CD_MEGA_PROCESSO, "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " ATUA_DT_ATUALIZACAO) VALUES ('"
   	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & strU & "'," & strMP & ", "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
   	
   	'response.write str_SQL_Nova_Sub_Empr
   	
	Set rdsNovo = conn_db.Execute(str_SQL_Novo_Acesso)

	strChave = CStr(strU) & " " & CStr(strMP) '& " " & CStr(strSP) & " " & CStr(strM) & " " & CStr(strA) & CStr(strT)
	
	''call grava_log(strChave,"ACESSO","I",0)
	
end sub

'guarda o conteúdo da String
str_valor = str_MegaSelecionado

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
	   call Grava_Novo_Acesso(str_Usuario, str_atual)
	   valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop

conn_db.Close
set conn_db = Nothing

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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="78%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp; </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de registros gravados :</font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=valor_total%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Relação
    entre Atividade x Empresa realizada com sucesso!</b></font> 
      
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="cadas_acesso.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Acesso</font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

if request("selUsuario") = "" then
   str_Usuario=0
else
   str_Usuario=Request("selUsuario")
end if
if Request("txtMegaSelecionado") = "" then
   str_MegaSelecionado = "0"
else
   str_MegaSelecionado = Request("txtMegaSelecionado")
end if   
'response.write str_Atividade

conn_db.execute("DELETE FROM " & Session("PREFIXO") & "ACESSO WHERE USUA_CD_USUARIO='" & str_Usuario & "'")
strChave = CStr(str_Usuario) '& CStr(str_Processo) &  CStr(str_SubProcesso) & CStr(str_Empresa) 
''call grava_log(strChave,"ACESSO","D",0)	

str_SQL_Novo_Acesso = ""
Sub Grava_Novo_Acesso(strU, strMP)
	str_SQL_Novo_Acesso = ""
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " INSERT INTO " & Session("PREFIXO") & "ACESSO ( "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " USUA_CD_USUARIO, MEPR_CD_MEGA_PROCESSO, "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & " ATUA_DT_ATUALIZACAO) VALUES ('"
   	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & strU & "'," & strMP & ", "
	str_SQL_Novo_Acesso = str_SQL_Novo_Acesso & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
   	
   	'response.write str_SQL_Nova_Sub_Empr
   	
	Set rdsNovo = conn_db.Execute(str_SQL_Novo_Acesso)

	strChave = CStr(strU) & " " & CStr(strMP) '& " " & CStr(strSP) & " " & CStr(strM) & " " & CStr(strA) & CStr(strT)
	
	''call grava_log(strChave,"ACESSO","I",0)
	
end sub

'guarda o conteúdo da String
str_valor = str_MegaSelecionado

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
	   call Grava_Novo_Acesso(str_Usuario, str_atual)
	   valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop

conn_db.Close
set conn_db = Nothing

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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="78%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp; </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de registros gravados :</font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=valor_total%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Relação
    entre Atividade x Empresa realizada com sucesso!</b></font> 
      
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="cadas_acesso.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Acesso</font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
