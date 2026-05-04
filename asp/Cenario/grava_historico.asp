<%@LANGUAGE="VBSCRIPT"%> 
 
<%


str_Cenario = Request("txtCenario")
str_DescHistorico = Request("txtDescHistorico")

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

'************************ INCLUSÃO DE OPERAÇÃO ESPECIAL **********************

   str_SQL_CenTran = ""
   str_SQL_CenTran = str_SQL_CenTran & " SELECT "
   str_SQL_CenTran = str_SQL_CenTran & " MAX(HICE_NR_SEQUENCIAL) AS MAX_SEQ "
   str_SQL_CenTran = str_SQL_CenTran & " FROM " & Session("PREFIXO") & "HISTORICO_CENARIO "
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

   ssql=" INSERT INTO " & Session("PREFIXO") & "HISTORICO_CENARIO ("
   ssql=ssql & " CENA_CD_CENARIO "
   ssql=ssql & " ,HICE_NR_SEQUENCIAL"
   ssql=ssql & " ,HICE_TX_HISTORICO"
   ssql=ssql & " ,ATUA_TX_OPERACAO "
   ssql=ssql & " ,ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO"
   ssql=ssql & " ) VALUES ( " 
   ssql=ssql & "'" & str_Cenario & "'"
   ssql=ssql & " ," & int_MaxCenTran 
   ssql=ssql & " ,'" & Ucase(str_DescHistorico) & "'"
   ssql=ssql & " ,'I', '" & Session("CdUsuario") & "', GETDATE())"

on error resume next
   
   'response.write ssql	
   
conn_db.execute(ssql)

'response.write err.description

if err.number=0 then
	erro=0
else
	erro=1
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
      <td height="20" width="43%">&nbsp; </td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="39%">&nbsp; </td>
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
      <td width="62%"> <font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Grava 
        hist&oacute;rico de cen&aacute;rio</font></td>
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
    <%if erro=0 then%>
    <tr> 
      <td width="3%"></td>
      <td width="24%"><%'=str_Processo%></td>
      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><b>Opera&ccedil;&atilde;o 
        realizada com Sucesso!</b></font></td>
      <td width="14%"></td>
    </tr>
    <% else %>
    <tr> 
      <td width="3%"></td>
      <td width="24%">&nbsp;</td>
      <td width="59%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#800000">Não 
        foi possível realizar a opera&ccedil;&atilde;o. Avise o problema.</font></b></td>
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
            <td height="41"><a href="javascript:history.go(-2)"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
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
