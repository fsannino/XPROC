 
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
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " DELETE FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " WHERE " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'" 
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " &  str_MegaProcesso2
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = " &  str_SubProcesso
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = '" & str_Transacao & "'" 
str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MODU_CD_MODULO = " & str_Modulo		
	
'response.write str_SQL_Nova_Ativ_Tran
Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ_Tran)

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" onload="javascript:window.close()">
exc 
<table width="22%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td><%=str_Funcao%></td>
  </tr>
  <tr> 
    <td><%=str_MegaProcesso%></td>
  </tr>
  <tr> 
    <td><%=str_MegaProcesso2%></td>
  </tr>
  <tr> 
    <td><%=str_Processo%></td>
  </tr>
  <tr> 
    <td><%=str_SubProcesso%></td>
  </tr>
  <tr> 
    <td><%=str_AtividadeCarga%></td>
  </tr>
  <tr> 
    <td><%=str_Modulo%></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
