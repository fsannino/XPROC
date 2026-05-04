<%
str_Opc = Request("txtOpc")

if (Request("selFuncao") <> "") then 
    str_Funcao = Request("selFuncao")
else
    str_Funcao = "0"
end if

if (Request("selMegaProcesso2") <> "") then 
    str_MegaProcesso2 = Request("selMegaProcesso2")
else
    str_MegaProcesso2 = "0"
end if

if (Request("selProcesso") <> "") then 
    str_Processo = Request("selProcesso")
else
    str_Processo = "0"
end if

if (Request("selSubProcesso") <> "") then 
    str_SubProcesso = Request("selSubProcesso")
else
    str_SubProcesso = "0"
end if

if (Request("selModulo") <> "") then 
    str_Modulo = Request("selModulo")
else
    str_Modulo = "0"
end if

if (Request("selAtividadeCarga") <> "") then 
    str_AtividadeCarga = Request("selAtividadeCarga")
else
    str_AtividadeCarga = "0"
end if

if (Request("selTransacao") <> "") then 
    str_Transacao = Request("selTransacao")
else
    str_Transacao = "0"
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao") 

str_SQL_Nova_Ativ_Tran = ""
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " INSERT INTO " & Session("PREFIXO") & "FUN_NEG_TRANSACAO ( "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " FUNE_CD_FUNCAO_NEGOCIO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,MEPR_CD_MEGA_PROCESSO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,PROC_CD_PROCESSO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,SUPR_CD_SUB_PROCESSO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,MODU_CD_MODULO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATCA_CD_ATIVIDADE_CARGA "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,TRAN_CD_TRANSACAO "	
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_TX_OPERACAO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_CD_NR_USUARIO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_DT_ATUALIZACAO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ) Values( "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & "'" & str_Funcao & "',"
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & str_MegaProcesso2 & "," & str_Processo & ","
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & str_SubProcesso & "," & str_Modulo & "," & str_AtividadeCarga & ","
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & "'" & str_Transacao & "'," & "'I', '" & Session("CdUsuario") & "', GETDATE())" 
	
Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ_Tran)

'response.redirect "http://its_server3/mail_resp_inc.asp?user=" & Session("CdUsuario") & "&mega=" & str_MegaProcesso2 & "&opt=inclusão&funcao="& str_funcao &"&transac=" & str_transacao & "&chave=" & session("CdUsuario")
response.redirect "mail_resp_inc.asp?user=" & Session("CdUsuario") & "&mega=" & str_MegaProcesso2 & "&opt=inclusão&funcao="& str_funcao &"&transac=" & str_transacao & "&chave=" & session("CdUsuario")
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" onload="javascript:window.close()">
</body>
</html>
