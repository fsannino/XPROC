<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../protege/protege.asp" -->
<%
str_Acao = Request("txtAcao")
str_CdCenario = UCase(Trim(Request("txtCdCenario")))
str_Cenario = Request("txtCenario")
str_CenarioChSequencia = Request("txtCenarioChSequencia")
str_CenarioTrSequencia  = Request("txtCenarioTrSequencia")
str_DescTransacao = Request("txtDescTransacao")
str_MegaProcesso2 = Request("selMegaProcesso2")
str_OpEs = Request("selOpEs")
str_BPP = Request("selBPP")
str_CenarioSeguinte = Request("selCenario")
str_Desenv = Request("selDesenv")
str_Situacao = Request("txtStNovo")

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Alterado = 0

'************************ ALTERACAO DE TRANSAÇÃO **********************
if str_Acao = "AT" then
   ssql=" UPDATE " & Session("PREFIXO") & "CENARIO_TRANSACAO SET "
   ssql=ssql & " CETR_TX_DESC_TRANSACAO = '" & Ucase(str_DescTransacao) & "'"
   ssql=ssql & " ,CENA_NR_SEQUENCIA_TRANS = " & str_CenarioTrSequencia
   ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'" 
   ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'" 
   ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   ssql=ssql & " AND CETR_NR_SEQUENCIA = " &  str_CenarioChSequencia 
   str_Alterado = 1
end if

'************************ INCLUSÃO DE EXIT/INTERFACE **********************
if str_Acao = "ED" then
   str_SQL_CenTran = ""
   str_SQL_CenTran = str_SQL_CenTran & " SELECT "
   str_SQL_CenTran = str_SQL_CenTran & " MAX(CETR_NR_SEQUENCIA) AS MAX_SEQ "
   str_SQL_CenTran = str_SQL_CenTran & " FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO "
   str_SQL_CenTran = str_SQL_CenTran & " GROUP BY CENA_CD_CENARIO "
   str_SQL_CenTran = str_SQL_CenTran & " HAVING CENA_CD_CENARIO = '" & str_Cenario & "'"

   'response.write str_SQL_CenTran	

   Set rdsMaxCenTran = Conn_db.Execute(str_SQL_CenTran)
	
   if rdsMaxCenTran.EOF then
      int_MaxCenTran = 1
   else
      int_MaxCenTran = rdsMaxCenTran("MAX_SEQ") + 1	
   end if
   rdsMaxCenTran.Close
   set rdsMaxCenTran = Nothing

   ssql=" INSERT INTO " & Session("PREFIXO") & "CENARIO_TRANSACAO ("
   ssql=ssql & " CENA_CD_CENARIO "
   ssql=ssql & " ,CETR_NR_SEQUENCIA"
   ssql=ssql & " ,OPES_CD_OPERACAO_ESP"
   ssql=ssql & " ,DESE_CD_DESENVOLVIMENTO"
   ssql=ssql & " ,CETR_TX_DESC_TRANSACAO"
   ssql=ssql & " ,CENA_NR_SEQUENCIA_TRANS"
   ssql=ssql & " ,CETR_TX_TIPO_RELACAO"
   ssql=ssql & " ,ATUA_TX_OPERACAO "
   ssql=ssql & " ,ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO"
   ssql=ssql & " ) VALUES ( " 
   ssql=ssql & "'" & str_Cenario & "'"
   ssql=ssql & " ," & int_MaxCenTran 
   ssql=ssql & " ," & str_OpEs
   ssql=ssql & " ,'" & str_Desenv & "'"
   ssql=ssql & " ,'" & Ucase(str_DescTransacao) & "'"
   ssql=ssql & " ," & str_CenarioTrSequencia 
   ssql=ssql & " , '3'" 
   ssql=ssql & " ,'I', '" & Session("CdUsuario") & "', GETDATE())"
   str_Alterado = 1
end if
'************************ ALTERACAO DE EXIT/INTERFACE **********************
if str_Acao = "AED" then
   ssql=" UPDATE " & Session("PREFIXO") & "CENARIO_TRANSACAO SET "
   ssql=ssql & " CETR_TX_DESC_TRANSACAO = '" & Ucase(str_DescTransacao) & "'"
   ssql=ssql & " ,CENA_NR_SEQUENCIA_TRANS = " & str_CenarioTrSequencia
   ssql=ssql & " ,OPES_CD_OPERACAO_ESP = " & str_OpEs
   ssql=ssql & " ,DESE_CD_DESENVOLVIMENTO = '" & str_Desenv & "'"
   ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'" 
   ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'" 
   ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   ssql=ssql & " AND CETR_NR_SEQUENCIA = " &  str_CenarioChSequencia 
   str_Alterado = 1
end if


'************************ INCLUSÃO DE OPERAÇÃO ESPECIAL **********************
if str_Acao = "IO" then
   str_SQL_CenTran = ""
   str_SQL_CenTran = str_SQL_CenTran & " SELECT "
   str_SQL_CenTran = str_SQL_CenTran & " MAX(CETR_NR_SEQUENCIA) AS MAX_SEQ "
   str_SQL_CenTran = str_SQL_CenTran & " FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO "
   str_SQL_CenTran = str_SQL_CenTran & " GROUP BY CENA_CD_CENARIO "
   str_SQL_CenTran = str_SQL_CenTran & " HAVING CENA_CD_CENARIO = '" & str_Cenario & "'"

   'response.write str_SQL_CenTran	

   Set rdsMaxCenTran = Conn_db.Execute(str_SQL_CenTran)
	
   if rdsMaxCenTran.EOF then
      int_MaxCenTran = 1
   else
      int_MaxCenTran = rdsMaxCenTran("MAX_SEQ") + 1	
   end if
   rdsMaxCenTran.Close
   set rdsMaxCenTran = Nothing

   ssql=" INSERT INTO " & Session("PREFIXO") & "CENARIO_TRANSACAO ("
   ssql=ssql & " CENA_CD_CENARIO "
   ssql=ssql & " ,CETR_NR_SEQUENCIA"
   ssql=ssql & " ,OPES_CD_OPERACAO_ESP"
   ssql=ssql & " ,CETR_TX_DESC_TRANSACAO"
   ssql=ssql & " ,CENA_NR_SEQUENCIA_TRANS"
   ssql=ssql & " ,CETR_TX_TIPO_RELACAO"
   ssql=ssql & " ,ATUA_TX_OPERACAO "
   ssql=ssql & " ,ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO"
   ssql=ssql & " ) VALUES ( " 
   ssql=ssql & "'" & str_Cenario & "'"
   ssql=ssql & " ," & int_MaxCenTran 
   ssql=ssql & " ," & str_OpEs
   ssql=ssql & " ,'" & Ucase(str_DescTransacao) & "'"
   ssql=ssql & " ," & str_CenarioTrSequencia 
   ssql=ssql & " , '1'" 
   ssql=ssql & " ,'I', '"& Session("CdUsuario") &"', GETDATE())"
   str_Alterado = 1
end if
'************************ ALTERACAO DE OPERACAO ESPECIAL **********************
if str_Acao = "AO" then
   ssql=" UPDATE " & Session("PREFIXO") & "CENARIO_TRANSACAO SET "
   ssql=ssql & " CETR_TX_DESC_TRANSACAO = '" & Ucase(str_DescTransacao) & "'"
   ssql=ssql & " ,CENA_NR_SEQUENCIA_TRANS = " & str_CenarioTrSequencia
   ssql=ssql & " ,OPES_CD_OPERACAO_ESP = " & str_OpEs
   ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'" 
   ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'" 
   ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   ssql=ssql & " AND CETR_NR_SEQUENCIA = " &  str_CenarioChSequencia 
   str_Alterado = 1
end if

'************************ INCLUSÃO DE CHAMADA CENARIO **********************
if str_Acao = "ICC" then

   if str_CdCenario <> "" then

      str_SQL = ""
      str_SQL = str_SQL & " SELECT "
      str_SQL = str_SQL & " CENA_CD_CENARIO, MEPR_CD_MEGA_PROCESSO "
      str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "CENARIO "
      str_SQL = str_SQL & " where CENA_CD_CENARIO = '" & str_CdCenario & "'"
      'response.write str_SQL	
      Set rdsCena = Conn_db.Execute(str_SQL)
	
      if not rdsCena.EOF then
         str_CenarioSeguinte = str_CdCenario
		 str_MegaProcesso2 = rdsCena("MEPR_CD_MEGA_PROCESSO")
	  else
		 str_CenarioSeguinte = ""
	  end if  
	       	  	
   end if
   if str_CenarioSeguinte <> ""	then
      str_SQL_CenTran = ""
      str_SQL_CenTran = str_SQL_CenTran & " SELECT "
      str_SQL_CenTran = str_SQL_CenTran & " MAX(CETR_NR_SEQUENCIA) AS MAX_SEQ "
      str_SQL_CenTran = str_SQL_CenTran & " FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO "
      str_SQL_CenTran = str_SQL_CenTran & " GROUP BY CENA_CD_CENARIO "
      str_SQL_CenTran = str_SQL_CenTran & " HAVING CENA_CD_CENARIO = '" & str_Cenario & "'"

      'response.write str_SQL_CenTran	

      Set rdsMaxCenTran = Conn_db.Execute(str_SQL_CenTran)
	
      if rdsMaxCenTran.EOF then
         int_MaxCenTran = 1
      else
         int_MaxCenTran = rdsMaxCenTran("MAX_SEQ") + 1	
      end if
      rdsMaxCenTran.Close
      set rdsMaxCenTran = Nothing

      ssql=" INSERT INTO " & Session("PREFIXO") & "CENARIO_TRANSACAO ("
      ssql=ssql & " CENA_CD_CENARIO "
      ssql=ssql & " ,CETR_NR_SEQUENCIA"
      ssql=ssql & " ,CENA_CD_CENARIO_SEGUINTE"
      ssql=ssql & " ,OPES_CD_OPERACAO_ESP"
      ssql=ssql & " ,CETR_TX_DESC_TRANSACAO"
      ssql=ssql & " ,CENA_NR_SEQUENCIA_TRANS"
      ssql=ssql & " ,MEPR_CD_MEGA_PROCESSO"   
      ssql=ssql & " ,CETR_TX_TIPO_RELACAO"
      ssql=ssql & " ,ATUA_TX_OPERACAO "
      ssql=ssql & " ,ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO"
      ssql=ssql & " ) VALUES ( " 
      ssql=ssql & "'" & str_Cenario & "'"
      ssql=ssql & " ," & int_MaxCenTran 
      if str_CenarioSeguinte <> "0" then
         ssql=ssql & " ,'" & str_CenarioSeguinte & "'" 
      else
         ssql=ssql & " , null "
      end if   
      ssql=ssql & " , 3 " 
      ssql=ssql & " ,'" & Ucase(str_DescTransacao) & "'"
      ssql=ssql & " ," & str_CenarioTrSequencia 
      ssql=ssql & " ," & str_MegaProcesso2 
      ssql=ssql & " ,'2'"
      ssql=ssql & " ,'I', '" & Session("CdUsuario") & "', GETDATE())"
      str_Alterado = 1
   else
      str_Alterado = 3	   
   end if 
end if

if str_Acao = "ACC" then
   ssql=" UPDATE " & Session("PREFIXO") & "CENARIO_TRANSACAO SET "
   ssql=ssql & " CETR_TX_DESC_TRANSACAO = '" & Ucase(str_DescTransacao) & "'"
   ssql=ssql & " ,CENA_NR_SEQUENCIA_TRANS = " & str_CenarioTrSequencia
   ssql=ssql & " ,MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso2
   ssql=ssql & " ,CENA_CD_CENARIO_SEGUINTE = '" & str_CenarioSeguinte & "'"
   ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'" 
   ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'" 
   ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   ssql=ssql & " AND CETR_NR_SEQUENCIA = " &  str_CenarioChSequencia 
   str_Alterado = 1
end if

if str_Acao = "ET" then
   ssql=" DELETE FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   ssql=ssql & " AND CETR_NR_SEQUENCIA = " &  str_CenarioChSequencia 
   str_Alterado = 1
end if

'************************ ALTERACAO situacao **********************
if str_Acao = "AS" then
   ssql=" UPDATE " & Session("PREFIXO") & "CENARIO SET "
   ssql=ssql & " CENA_TX_SITUACAO = '" & str_Situacao & "'"
   ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'" 
   ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'" 
   ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   str_Alterado = 1
end if

'response.write ssql
'***************************** BPP *****************************

if str_Acao = "BPP" then
   ssql=" UPDATE " & Session("PREFIXO") & "CENARIO_TRANSACAO SET "
   if str_BPP <> "0" then
      ssql=ssql & " BPPP_CD_BPP = '" & Ucase(str_BPP) & "'"
   else
      ssql=ssql & " BPPP_CD_BPP = NULL "
   end if      
   ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'" 
   ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" &  Session("CdUsuario") & "'"
   ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   ssql=ssql & " AND CETR_NR_SEQUENCIA = " &  str_CenarioChSequencia 
   str_Alterado = 1
end if

'***************************************************************

if str_Acao = "XX" then

   ssql=" UPDATE " & Session("PREFIXO") & "CENARIO_TRANSACAO SET "
   if str_BPP <> "0" then
      ssql=ssql & " BPPP_CD_BPP = " & str_BPP
   else
      ssql=ssql & " BPPP_CD_BPP = null "
   end if	  
   if str_CenarioSeguinte <> "0" then
      ssql=ssql & " ,CENA_CD_CENARIO_SEGUINTE = " & str_CenarioSeguinte
   else
      ssql=ssql & " ,CENA_CD_CENARIO_SEGUINTE = null "   
   end if	  
   if str_OpEs <> "0" then
      ssql=ssql & " ,OPES_CD_OPERACAO_ESP = " & str_OpEs
   else
      ssql=ssql & " ,OPES_CD_OPERACAO_ESP = null "
   end if   	  
   ssql=ssql & " ,CETR_TX_DESC_TRANSACAO = '" & Ucase(str_DescOprEs) & "'"
   ssql=ssql & " ,CENA_NR_SEQUENCIA_TRANS = " & str_Seq
   ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'" 
   ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'" 
   ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   ssql=ssql & " AND CETR_NR_SEQUENCIA = " &  str_Cen_Seq 
end if

'on error resume next

'response.write ssql
if str_Alterado <> 3 then
   conn_db.execute(ssql)
end if
'response.write err.description

'response.write str_Alterado

if str_Alterado = 1 then

   str_SQL_CenStatus = ""
   str_SQL_CenStatus = str_SQL_CenStatus & " SELECT "
   str_SQL_CenStatus = str_SQL_CenStatus & " CENA_TX_SITUACAO_LOTUS "
   str_SQL_CenStatus = str_SQL_CenStatus & " FROM " & Session("PREFIXO") & "CENARIO "
   str_SQL_CenStatus = str_SQL_CenStatus & " WHERE CENA_CD_CENARIO = '" & str_Cenario & "'"

   'response.write str_SQL_CenStatus	

   Set rdsStatusLotus = Conn_db.Execute(str_SQL_CenStatus)
	
   if NOT rdsStatusLotus.EOF then
      int_StatusLotus = rdsStatusLotus("CENA_TX_SITUACAO_LOTUS")
	  AA =  rdsStatusLotus("CENA_TX_SITUACAO_LOTUS")
   else
   	  int_StatusLotus = "NC"	
   end if
   rdsStatusLotus.Close
   set rdsStatusLotus = Nothing
   
   select case int_StatusLotus
      case "NC"
	   int_StatusLotus="NC"
      case "CR"
	   int_StatusLotus="RC"
      case "RC"
	   int_StatusLotus="RC"
      case else
	   int_StatusLotus="NC"
   end select
   
   ssql=" UPDATE " & Session("PREFIXO") & "CENARIO SET "
   ssql=ssql & " CENA_TX_SITUACAO_LOTUS = '" & int_StatusLotus & "'"
   ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'" 
   ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
   ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE() "
   ssql=ssql & " where CENA_CD_CENARIO = '" & str_Cenario & "'"
   conn_db.execute(ssql)
   'RESPONSE.WRITE ssql
end if

if err.number=0 then
	erro=0
else
	erro=1
end if

if str_Alterado = 3 then
   erro = 3
end if

str_Sql_Cenario = ""
str_Sql_Cenario = str_Sql_Cenario & " SELECT "
str_Sql_Cenario = str_Sql_Cenario & " " & Session("PREFIXO") & "CENARIO.MEPR_CD_MEGA_PROCESSO, "
str_Sql_Cenario = str_Sql_Cenario & " " & Session("PREFIXO") & "CENARIO.PROC_CD_PROCESSO, "
str_Sql_Cenario = str_Sql_Cenario & " " & Session("PREFIXO") & "CENARIO.SUPR_CD_SUB_PROCESSO "
str_Sql_Cenario = str_Sql_Cenario & " FROM "
str_Sql_Cenario = str_Sql_Cenario & " " & Session("PREFIXO") & "CENARIO "
str_Sql_Cenario = str_Sql_Cenario & " WHERE "
str_Sql_Cenario = str_Sql_Cenario & " " & Session("PREFIXO") & "CENARIO.CENA_CD_CENARIO = '" & str_Cenario & "'"

Set rdsCenario= Conn_db.Execute(str_Sql_Cenario)
if not rdsCenario.EOF then
   str_MegaProcesso = rdsCenario("MEPR_CD_MEGA_PROCESSO")
   str_Processo = rdsCenario("PROC_CD_PROCESSO")
   str_SubProcesso = rdsCenario("SUPR_CD_SUB_PROCESSO")   
else
   str_MegaProcesso = "0"
   str_Processo = "0"
   str_SubProcesso = "0"
end if
rdsCenario.Close  

set rdsCenario = Nothing
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
<form method="POST" action="../valida_altera_processo.asp?mega=<%=str_MegaProcesso%>&Proc=<%=str_Processo%>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr>
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="3%">&nbsp;</td>
      <td height="20" width="43%"><%'=str_Alterado%> </td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="39%">&nbsp; </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"><%'=str_CdCenario%> - <%'=str_Cenario%> - <%'=str_Alterado%></td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%">
	  <% if str_Acao = "ET" then
	        ls_Titulo = "Exclui transação de cenário."
	     else
  	        ls_Titulo = "Grava transação de cenário."
	     end if
	  %>
        <font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><%=ls_Titulo%></font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_MegaProcesso%></td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <% if erro=0 then %>
    <tr> 
      <td width="3%"></td>
      <td width="24%"><%'=str_Processo%></td>
      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><b>Opera&ccedil;&atilde;o 
        realizada com Sucesso!</b></font></td>
      <td width="14%"></td>
    </tr>
    <% elseif erro = 1 then %>
    <tr> 
      <td width="3%"></td>
      <td width="24%">&nbsp;</td>
      <td width="59%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#800000">Não 
        foi possível realizar a opera&ccedil;&atilde;o. Avise o problema.</font></b></td>
      <td width="14%"></td>
    </tr>
    <% elseif erro = 3 then  %>
    <tr> 
      <td width="3%"></td>
      <td width="24%">&nbsp;</td>
      <td width="59%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#800000">Não 
        foi possível realizar a opera&ccedil;&atilde;o. Cen&aacute;rio n&atilde;o 
        encontrado.</font></b></td>
      <td width="14%"></td>
    </tr>	
    <% end if %>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_SubProcesso%></td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_Cenario%></td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">
        <table width="74%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr> 
            <td height="41"><a href="gerencia_cenario_transa.asp?selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>&selCenario=<%=str_Cenario%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela de Edi&ccedil;&atilde;o de Cen&aacute;rio</font></td>
          </tr>
          <tr> 
            <td height="41"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
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
</body>
</html>
