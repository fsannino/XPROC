<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conecta.asp" -->
<%
set objUSR = server.createobject("Seseg.Usuario")

if objUSR.GetUsuario then
      	Usuario = objUSR.sei_chave
end if

set objUSR=Nothing

server.scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

mega = request("selMega")
assunto = request("selModulo")

org1 = request("str01")
org2 = request("str02")
org3 = request("str03")

if org3<>0 then
	orgao = org3 & "00000"
else
	if org2<>0 then
		orgao=org2 & "00000000"
	else
		orgao = org1
	end if
end if

titulo = UCASE(request("txttitulo"))
descricao = UCASE(request("txtdescricao"))
solicitante = UCASE(request("txtsolicitante"))

chave = ucase(request("txtchave"))
fone = ucase(request("txtfone"))

responsavel = request("selResponsavel")
prioridade = request("selPrioridade")
tipo = request("selTipo")
legado = request("selLegado")

set rs = db.execute("SELECT MAX(BALO_CD_COD_BACKLOG) AS CODIGO FROM BACKLOG")


if not isnull(rs("CODIGO")) then
	cod_reg = rs("CODIGO") + 1
else
	cod_reg = 1
end if

ssql=""
ssql="INSERT INTO BACKLOG(BALO_CD_COD_BACKLOG, MEPR_CD_MEGA_PROCESSO, SUMO_NR_CD_SEQUENCIA, ORME_CD_ORG_MENOR, BALO_TX_TITULO, BALO_TX_DESCRICAO, BALO_TX_SOLICITANTE, BALO_CD_RESPONSAVEL, BALO_CD_PRIORIDADE, BALO_CD_TIPO, BALO_CD_LEGADO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, BALO_TX_CHAVE, BALO_TX_TELEFONE)"
ssql=ssql+"VALUES(" & cod_reg & "," & mega & "," & assunto & ",'" & orgao & "','" & left(titulo,99) & "','" & left(descricao,999) & "','" & left(solicitante,49) & "'," & responsavel & ","& prioridade &"," & tipo & "," & legado & ",'I','" & Usuario & "',GETDATE(),'" & chave & "','" & fone & "')"

db.execute(ssql)

%>
<html>
<head>
<title>#BACKLOG - Solicitações de Melhoria no SAP R/3#</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#000099" alink="#000099" link="#000099">
<p>&nbsp;</p>
<p align="center"><font face="Verdana" color="#000080">Solicita&ccedil;&otilde;es 
  de Melhoria na Solu&ccedil;&atilde;o Configurada no SAP R/3</font> </p>
<p align="center">&nbsp;</p>
<p align="center"><font face="Verdana" color="#000080"><b><font size="2">O Registro 
  foi gravado com Sucesso!</font></b></font></p>
<table width="75%" border="0" align="center">
  <tr> 
    <td width="30%" height="51">&nbsp;</td>
    <td width="6%" height="51"> 
      <div align="center"><img src="seta_d.jpg" width="23" height="24"></div>
    </td>
    <td width="64%" height="51"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099"><a href="cad_backlog.asp">Retornar 
      ao Cadastro de Solicita&ccedil;&atilde;o</a></font></td>
  </tr>
  <tr> 
    <td width="30%" height="58">&nbsp;</td>
    <td width="6%" height="58"> 
      <div align="center"><img src="seta_d.jpg" width="23" height="24"></div>
    </td>
    <td width="64%" height="58"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099"><a href="index.asp">Retornar 
      ao Menu Principal</a></font></td>
  </tr>
</table>
<p align="center">&nbsp;</p>
</body>
</html>