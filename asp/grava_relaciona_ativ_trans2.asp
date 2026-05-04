<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->

<%
Dim str_Opc
dim str_MegaProcesso
dim str_Processo
dim str_SubProcesso
dim str_AtividadeCarga
dim str_Modulo
dim str_Transacao

dim int_Total_Atividade
dim int_Total_AtividadeTrans

str_Opc = Request("txtOpc")

str_MegaProcesso= Request("selMegaProcesso")
str_Processo = Request("selProcesso")
str_SubProcesso = Request("selSubProcesso")
str_AtividadeCarga = Request("selAtividadeCarga")
str_Modulo = Request("selModulo")

str_DescMegaProcesso= Request("txtDsMP")
str_DescProcesso = Request("txtDsP")
str_DescSubProcesso = Request("txtDsSP")
str_DescAtividadeCarga = Request("txtDsA")
str_DescModulo = Request("txtDsM")
str_Tran_list2 = Request("list2")
str_Transacao = Request("txtTranSelecionada")
str_DsMP = Request("txtDsMP")

int_Total_Atividade = 0
int_Total_AtividadeTrans = 0

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

Sub Grava_Nova_Atividade(strMP, strP, strSP, strA)

    str_SQL_Nova_Ativ = ""
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " INSERT INTO " & Session("PREFIXO") & "ATIVIDADE ( "
    str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,PROC_CD_PROCESSO "
    str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,SUPR_CD_SUB_PROCESSO "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,ATIV_CD_ATIVIDADE "
    str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,ATUA_TX_OPERACAO "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ) Values( "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & strMP & "," & strP & ","
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & strSP & "," & strA & ","
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " 'I', 'XXXX', GETDATE())" 
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ)
	strChave = CStr(strMP) & CStr(strP) & CStr(strSP) & CStr(strA)
	'call grava_log(strChave,"ATIVIDADE","I",0)
	
    int_Total_Atividade = int_Total_Atividade + 1
	
end sub

Sub Grava_Nova_Atividade_Trans(strMP, strP, strSP, strA, strT)

    str_SQL_Nova_Ativ_Tran = ""
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " INSERT INTO " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO ( "
    str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,PROC_CD_PROCESSO "
    str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,SUPR_CD_SUB_PROCESSO "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATIV_CD_ATIVIDADE "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,TRAN_CD_TRANSACAO "
    str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_TX_OPERACAO "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ) Values( "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strMP & "," & strP & ","
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strSP & "," & strA & ","
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & "'" & strT & "'," & " 'I', 'XXXX', GETDATE())" 
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ_Tran)

	strChave = CStr(strMP) & " " & CStr(strP) & " " &  CStr(strSP) & " " &  CStr(strA) & " " &  CStr(strT)
	'call grava_log(strChave,"ATIVIDADE","I",0)
	
    int_Total_AtividadeTrans = int_Total_AtividadeTrans + 1
	
end sub

str_SQL_Deleta_Ativ_Tran = ""
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " DELETE " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " FROM " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " INNER JOIN " & Session("PREFIXO") & "TRANSACAO ON " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " INNER JOIN " & Session("PREFIXO") & "MODULO_R3 ON " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " WHERE "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO = '" & str_Modulo & "'"

Set rdsDeleta = conn_db.Execute(str_SQL_Deleta_Ativ_Tran)

str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " SELECT " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " FROM " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " INNER JOIN " & Session("PREFIXO") & "TRANSACAO ON " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " INNER JOIN " & Session("PREFIXO") & "MODULO_R3 ON " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " WHERE "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " AND " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO <> '" & str_Modulo & "'"

Set rdsExistem_Mais_Ativ_Transacoes = conn_db.Execute(str_SQL_Existe_Mais_Tran)

if rdsExistem_Mais_Ativ_Transacoes.EOF Then
   str_SQL_Deleta_Ativ = ""
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " Delete from " & Session("PREFIXO") & "ATIVIDADE "
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " WHERE "
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " " & Session("PREFIXO") & "ATIVIDADE.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " AND " & Session("PREFIXO") & "ATIVIDADE.PROC_CD_PROCESSO = " & str_Processo
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " AND " & Session("PREFIXO") & "ATIVIDADE.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " AND " & Session("PREFIXO") & "ATIVIDADE.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga

   Set rdsDeleta = conn_db.Execute(str_SQL_Deleta_Ativ)
   str_Existem_Mais_Ativ_Trans = "N"
else
   str_Existem_Mais_Ativ_Trans = "S"
end if

if str_Existem_Mais_Ativ_Trans = "N" then
   call Grava_Nova_Atividade(str_MegaProcesso, str_Processo, str_SubProcesso, str_AtividadeCarga)
end if

int_Tamanho = Len(Trim(str_Transacao))
str_Trata = Trim(Mid(str_Transacao,2,int_Tamanho))
'str_Transacao = str_Trata
for i=1 to int_Tamanho
    if Mid(str_Trata,i,1) = "["  then
       str_NovaTransacao = Trim(Mid(str_Trata,1,i-1))
       str_Trata = Trim(Mid(str_Trata,i+1,int_Tamanho))
	   call Grava_Nova_Atividade_Trans(str_MegaProcesso, str_Processo, str_SubProcesso, str_AtividadeCarga, str_NovaTransacao)
    end if
next
call Grava_Nova_Atividade_Trans(str_MegaProcesso, str_Processo, str_SubProcesso, str_AtividadeCarga, str_Trata)

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
<table width="83%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td width="20%"> 
      <%'=str_Opc%>
      <%'=str_MegaProcesso%>
      <%'=str_Processo%>
      <%'=str_SubProcesso%>
    </td>
    <td width="6%">&nbsp;</td>
    <td width="62%"> 
      <%'=str_AtividadeCarga%>
      <%'=str_Modulo%>
      <%'=str_Transacao%>
      <%'=i%>
      <%'=str_NovaTransacao%>
      <%'=str_Trata%>
      <%'=int_Total_Atividade%>
    </td>
    <td width="12%"></td>
    <td width="12%"> 
      <%'=str_Tran_list2%>
    </td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade</font><font face="Arial, Helvetica, sans-serif" size="2">: 
        &nbsp;</font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_AtividadeCarga%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescAtividadeCarga%></font></font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Mega 
        Processo:&nbsp; </font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_MegaProcesso%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      &nbsp;&nbsp;- <%=str_DescMegaProcesso%></font></font></td>
    <td width="12%">&nbsp; </td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Processo: 
        &nbsp;</font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_Processo%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescProcesso%></font></font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Sub 
        Processo: &nbsp;</font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_SubProcesso%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescSubProcesso%></font></font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">M&oacute;dulo</font><font face="Arial, Helvetica, sans-serif" size="2">: 
        &nbsp;</font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_Modulo%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescModulo%></font></font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"></font></div>
    </td>
    <td width="6%">&nbsp;</td>
    <td width="62%"><font color="#003366"> </font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr bgcolor="#0099CC"> 
    <td width="20%" height="7"></td>
    <td width="6%" height="7"></td>
    <td width="62%" height="7"></td>
    <td width="12%" height="7"></td>
    <td width="12%" height="7"></td>
  </tr>
</table>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de transa&ccedil;&otilde;es gravadas:<%=int_Total_AtividadeTrans%> </font></td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%"> 
      <table width="98%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=2&selAtividadeCarga=0&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona nova Atividade" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            nova Atividade</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=9&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=0&selProcesso=0&selSubProcesso=0"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Mega Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=9&selModulo=<%=str_Modulo%>&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=0&selSubProcesso=0""><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=9&selModulo=<%=str_Modulo%>&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=0"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Selecuina 
            novo Sub Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=3&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo M&oacute;dulo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo M&oacute;dulo</font></td>
        </tr>
      </table>
    </td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p> 
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->

<%
Dim str_Opc
dim str_MegaProcesso
dim str_Processo
dim str_SubProcesso
dim str_AtividadeCarga
dim str_Modulo
dim str_Transacao

dim int_Total_Atividade
dim int_Total_AtividadeTrans

str_Opc = Request("txtOpc")

str_MegaProcesso= Request("selMegaProcesso")
str_Processo = Request("selProcesso")
str_SubProcesso = Request("selSubProcesso")
str_AtividadeCarga = Request("selAtividadeCarga")
str_Modulo = Request("selModulo")

str_DescMegaProcesso= Request("txtDsMP")
str_DescProcesso = Request("txtDsP")
str_DescSubProcesso = Request("txtDsSP")
str_DescAtividadeCarga = Request("txtDsA")
str_DescModulo = Request("txtDsM")
str_Tran_list2 = Request("list2")
str_Transacao = Request("txtTranSelecionada")
str_DsMP = Request("txtDsMP")

int_Total_Atividade = 0
int_Total_AtividadeTrans = 0

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

Sub Grava_Nova_Atividade(strMP, strP, strSP, strA)

    str_SQL_Nova_Ativ = ""
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " INSERT INTO " & Session("PREFIXO") & "ATIVIDADE ( "
    str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,PROC_CD_PROCESSO "
    str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,SUPR_CD_SUB_PROCESSO "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,ATIV_CD_ATIVIDADE "
    str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,ATUA_TX_OPERACAO "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " ) Values( "
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & strMP & "," & strP & ","
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & strSP & "," & strA & ","
	str_SQL_Nova_Ativ = str_SQL_Nova_Ativ & " 'I', 'XXXX', GETDATE())" 
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ)
	strChave = CStr(strMP) & CStr(strP) & CStr(strSP) & CStr(strA)
	'call grava_log(strChave,"ATIVIDADE","I",0)
	
    int_Total_Atividade = int_Total_Atividade + 1
	
end sub

Sub Grava_Nova_Atividade_Trans(strMP, strP, strSP, strA, strT)

    str_SQL_Nova_Ativ_Tran = ""
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " INSERT INTO " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO ( "
    str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,PROC_CD_PROCESSO "
    str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,SUPR_CD_SUB_PROCESSO "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATIV_CD_ATIVIDADE "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,TRAN_CD_TRANSACAO "
    str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_TX_OPERACAO "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ) Values( "
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strMP & "," & strP & ","
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strSP & "," & strA & ","
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & "'" & strT & "'," & " 'I', 'XXXX', GETDATE())" 
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ_Tran)

	strChave = CStr(strMP) & " " & CStr(strP) & " " &  CStr(strSP) & " " &  CStr(strA) & " " &  CStr(strT)
	'call grava_log(strChave,"ATIVIDADE","I",0)
	
    int_Total_AtividadeTrans = int_Total_AtividadeTrans + 1
	
end sub

str_SQL_Deleta_Ativ_Tran = ""
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " DELETE " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " FROM " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " INNER JOIN " & Session("PREFIXO") & "TRANSACAO ON " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " INNER JOIN " & Session("PREFIXO") & "MODULO_R3 ON " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " WHERE "
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga
str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO = '" & str_Modulo & "'"

Set rdsDeleta = conn_db.Execute(str_SQL_Deleta_Ativ_Tran)

str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " SELECT " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " FROM " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " INNER JOIN " & Session("PREFIXO") & "TRANSACAO ON " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " INNER JOIN " & Session("PREFIXO") & "MODULO_R3 ON " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " WHERE "
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga
str_SQL_Existe_Mais_Tran = str_SQL_Existe_Mais_Tran & " AND " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO <> '" & str_Modulo & "'"

Set rdsExistem_Mais_Ativ_Transacoes = conn_db.Execute(str_SQL_Existe_Mais_Tran)

if rdsExistem_Mais_Ativ_Transacoes.EOF Then
   str_SQL_Deleta_Ativ = ""
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " Delete from " & Session("PREFIXO") & "ATIVIDADE "
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " WHERE "
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " " & Session("PREFIXO") & "ATIVIDADE.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " AND " & Session("PREFIXO") & "ATIVIDADE.PROC_CD_PROCESSO = " & str_Processo
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " AND " & Session("PREFIXO") & "ATIVIDADE.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
   str_SQL_Deleta_Ativ = str_SQL_Deleta_Ativ & " AND " & Session("PREFIXO") & "ATIVIDADE.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga

   Set rdsDeleta = conn_db.Execute(str_SQL_Deleta_Ativ)
   str_Existem_Mais_Ativ_Trans = "N"
else
   str_Existem_Mais_Ativ_Trans = "S"
end if

if str_Existem_Mais_Ativ_Trans = "N" then
   call Grava_Nova_Atividade(str_MegaProcesso, str_Processo, str_SubProcesso, str_AtividadeCarga)
end if

int_Tamanho = Len(Trim(str_Transacao))
str_Trata = Trim(Mid(str_Transacao,2,int_Tamanho))
'str_Transacao = str_Trata
for i=1 to int_Tamanho
    if Mid(str_Trata,i,1) = "["  then
       str_NovaTransacao = Trim(Mid(str_Trata,1,i-1))
       str_Trata = Trim(Mid(str_Trata,i+1,int_Tamanho))
	   call Grava_Nova_Atividade_Trans(str_MegaProcesso, str_Processo, str_SubProcesso, str_AtividadeCarga, str_NovaTransacao)
    end if
next
call Grava_Nova_Atividade_Trans(str_MegaProcesso, str_Processo, str_SubProcesso, str_AtividadeCarga, str_Trata)

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
<table width="83%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td width="20%"> 
      <%'=str_Opc%>
      <%'=str_MegaProcesso%>
      <%'=str_Processo%>
      <%'=str_SubProcesso%>
    </td>
    <td width="6%">&nbsp;</td>
    <td width="62%"> 
      <%'=str_AtividadeCarga%>
      <%'=str_Modulo%>
      <%'=str_Transacao%>
      <%'=i%>
      <%'=str_NovaTransacao%>
      <%'=str_Trata%>
      <%'=int_Total_Atividade%>
    </td>
    <td width="12%"></td>
    <td width="12%"> 
      <%'=str_Tran_list2%>
    </td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade</font><font face="Arial, Helvetica, sans-serif" size="2">: 
        &nbsp;</font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_AtividadeCarga%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescAtividadeCarga%></font></font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Mega 
        Processo:&nbsp; </font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_MegaProcesso%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      &nbsp;&nbsp;- <%=str_DescMegaProcesso%></font></font></td>
    <td width="12%">&nbsp; </td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Processo: 
        &nbsp;</font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_Processo%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescProcesso%></font></font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Sub 
        Processo: &nbsp;</font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_SubProcesso%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescSubProcesso%></font></font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">M&oacute;dulo</font><font face="Arial, Helvetica, sans-serif" size="2">: 
        &nbsp;</font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_Modulo%></font> </font></div>
    </td>
    <td width="62%"><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescModulo%></font></font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"> 
      <div align="right"><font color="#003366"></font></div>
    </td>
    <td width="6%">&nbsp;</td>
    <td width="62%"><font color="#003366"> </font></td>
    <td width="12%">&nbsp;</td>
    <td width="12%">&nbsp;</td>
  </tr>
  <tr bgcolor="#0099CC"> 
    <td width="20%" height="7"></td>
    <td width="6%" height="7"></td>
    <td width="62%" height="7"></td>
    <td width="12%" height="7"></td>
    <td width="12%" height="7"></td>
  </tr>
</table>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de transa&ccedil;&otilde;es gravadas:<%=int_Total_AtividadeTrans%> </font></td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%"> 
      <table width="98%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=2&selAtividadeCarga=0&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona nova Atividade" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            nova Atividade</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=9&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=0&selProcesso=0&selSubProcesso=0"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Mega Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=9&selModulo=<%=str_Modulo%>&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=0&selSubProcesso=0""><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=9&selModulo=<%=str_Modulo%>&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=0"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Selecuina 
            novo Sub Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans4.asp?txtOpc=3&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo M&oacute;dulo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo M&oacute;dulo</font></td>
        </tr>
      </table>
    </td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%">&nbsp;</td>
    <td width="34%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p> 
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
