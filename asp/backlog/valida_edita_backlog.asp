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

cod_reg = Session("Registro_atual")

ssql=""
ssql="UPDATE BACKLOG "
ssql=ssql+ "SET MEPR_CD_MEGA_PROCESSO=" & mega & ", "
ssql=ssql+ "SUMO_NR_CD_SEQUENCIA=" & assunto & ", "
ssql=ssql+ "ORME_CD_ORG_MENOR='" & orgao & "', "
ssql=ssql+ "BALO_TX_TITULO='" & left(titulo,99) & "', "
ssql=ssql+ "BALO_TX_DESCRICAO='" & left(descricao,999) & "', "
ssql=ssql+ "BALO_TX_SOLICITANTE='" & left(solicitante,49) & "', "
ssql=ssql+ "BALO_CD_RESPONSAVEL=" & responsavel & ", "
ssql=ssql+ "BALO_CD_PRIORIDADE="& prioridade &", "
ssql=ssql+ "BALO_CD_TIPO=" & tipo & ", "
ssql=ssql+ "BALO_CD_LEGADO=" & legado & ", "
ssql=ssql+ "ATUA_TX_OPERACAO='A', "
ssql=ssql+ "ATUA_CD_NR_USUARIO='" & Usuario & "', "
ssql=ssql+ "ATUA_DT_ATUALIZACAO=GETDATE(), "
ssql=ssql+ "BALO_TX_CHAVE='" & chave & "', " 
ssql=ssql+ "BALO_TX_TELEFONE= '" & fone & "'"
ssql=ssql+ " WHERE BALO_CD_COD_BACKLOG=" & cod_reg

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
  foi Alterado com Sucesso!</font></b></font></p>
<table width="75%" border="0" align="center">
  <tr> 
    <td width="30%" height="51">&nbsp;</td>
    <td width="6%" height="51"> 
      <div align="center"><img src="seta_d.jpg" width="23" height="24"></div>
    </td>
    <td width="64%" height="51"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000099"><a href="consulta_backlog.asp">Editar 
      outro Registro</a></font></td>
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