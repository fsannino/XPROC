<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->

<%
tem_erro = 0 

Dim str_Opc
dim str_MegaProcesso
dim str_Processo
dim str_SubProcesso
dim str_AtividadeCarga
dim str_Modulo
dim str_Transacao

dim int_Total_Atividade
dim int_Total_AtividadeTrans
Dim int_Tot_Trans_Exc
int_Tot_Trans_Exc = 0

str_Opc = Request("txtOpc")

str_MegaProcesso= Request("txtMP")
str_Processo = Request("txtP")
str_SubProcesso = Request("txtSP")
str_AtividadeCarga = Request("selAtividadeCarga")
str_Modulo = Request("selModulo")

str_DescMegaProcesso= Request("txtDsMP")
str_DescProcesso = Request("txtDsP")
str_DescSubProcesso = Request("txtDsSP")
str_DescAtividadeCarga = Request("txtDsA")
str_DescModulo = Request("txtDsM")

str_Transacao = Request("txtTranSelecionada")
str_NaoTransacao = Request("txtTranNaoSelecionada")
'response.Write "  Selec   "
'response.Write(str_Transacao)
'response.Write "  Nao Selec   "
'response.Write(str_NaoTransacao)

str_DsMP = Request("txtDsMP")

int_Total_Atividade = 0
int_Total_AtividadeTrans = 0

dim ls_Trans_Exist(50)
dim ls_Mega_Exist(50)
ls_Indice = 0
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

ls_Sequencia = 0

Sub Grava_Nova_Atividade_Trans(strMP, strP, strSP, strM, strA, strT)
    
	'str_SQL = ""
	'str_SQL = str_SQL & " SELECT TRAN_CD_TRANSACAO, MEPR_TX_DESC_MEGA_PROCESSO "
	'str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "RELACAO_FINAL, " & Session("PREFIXO") & "MEGA_PROCESSO "
	'str_SQL = str_SQL & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & strT & "'"
	'str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO <> " & strMP
	'str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
	''RESPONSE.WRITE str_SQL
	'Set rdsExiste = conn_db.Execute(str_SQL)
	'if not rdsExiste.EOF then
	'   ls_Trans_Exist(ls_Indice) = rdsExiste("TRAN_CD_TRANSACAO")
	'   ls_Mega_Exist(ls_Indice) = rdsExiste("MEPR_TX_DESC_MEGA_PROCESSO")
	'   ls_Indice = ls_Indice + 1
	'else
       ls_Sequencia = ls_Sequencia + 1
       str_SQL_Nova_Ativ_Tran = ""
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " INSERT INTO " & Session("PREFIXO") & "RELACAO_FINAL ( "
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " MEPR_CD_MEGA_PROCESSO "
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,PROC_CD_PROCESSO "
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,SUPR_CD_SUB_PROCESSO "
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,MODU_CD_MODULO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATCA_CD_ATIVIDADE_CARGA "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,TRAN_CD_TRANSACAO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,RELA_NR_SEQUENCIA "	
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_TX_OPERACAO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_CD_NR_USUARIO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_DT_ATUALIZACAO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ) Values( "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strMP & "," & strP & ","
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strSP & "," & strM & "," & strA & ","
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & "'" & strT & "'," & ls_Sequencia  & ", 'I', '" & Session("CdUsuario") & "', GETDATE())" 
	   Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ_Tran)
	   strChave = CStr(strMP) & " " & CStr(strP) & " " & CStr(strSP) & " " & CStr(strM) & " " & CStr(strA) & CStr(strT)
	   'call grava_log(strChave,"RELACAO_FINAL","I",0)
       int_Total_AtividadeTrans = int_Total_AtividadeTrans + 1
	'end if
end sub

Function Existe_Atividade_Trans(pMegaProcesso, pProcesso, pSubProcesso, pModulo, pAtividadeCarga, pAtual)
   str_SQL_Deleta_Ativ_Tran = ""
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " Select *  "
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " FROM " & Session("PREFIXO") & "RELACAO_FINAL "
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " WHERE "
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & pMegaProcesso
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & pProcesso
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & pSubProcesso
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MODU_CD_MODULO = " & pModulo 
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & pAtividadeCarga
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & pAtual & "'"
   'response.write str_SQL_Deleta_Ativ_Tran
   Set rdsExiste = conn_db.Execute(str_SQL_Deleta_Ativ_Tran)
   if rdsExiste.EOF then
      Existe_Atividade_Trans = False
   else	  
      Existe_Atividade_Trans = True
   end if
end function

'guarda o conteúdo da String
str_valor = str_Transacao

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
if tamanho > 0 then
'início da Rotina
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        if Existe_Atividade_Trans(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual) = false then
           'Aqui entra o que vc quer fazer com o caracter em questão!
	       call Grava_Nova_Atividade_Trans(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual)
        end if
        quantos = 0
    End If
    contador = contador + 1
Loop
end if


'guarda o conteúdo da String
str_valor = str_NaoTransacao
'response.Write(str_valor)
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
if tamanho > 0 then
'início da Rotina
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        if Existe_Relac(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual) = False then
           If Existe_em_Rel_Final(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual) = True then
              'Aqui entra o que vc quer fazer com o caracter em questão!
	          call Grava_Delecao_Transa(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual)
           end if
		end if
        quantos = 0
    End If
    contador = contador + 1
Loop
end if

Function Existe_em_Rel_Final(strMP, strP, strSP, strM, strA, strT)
	str_SQL = ""
	str_SQL = str_SQL & " Select *  from " & Session("PREFIXO") & "RELACAO_FINAL "
    str_SQL = str_SQL & " WHERE "
    str_SQL = str_SQL & " " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strMP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & strP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & strSP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MODU_CD_MODULO = " & strM 
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & strA
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & strT & "'"
    Set rsExiste_em_Rel_Final = conn_db.Execute(str_SQL)
	'response.Write(str_SQL)
    'response.write str_SQL
    Set rsExiste_em_Rel_Final = conn_db.Execute(str_SQL)
    if rsExiste_em_Rel_Final.EOF then
       Existe_em_Rel_Final = False
    else	  
       Existe_em_Rel_Final = True
    end if
end function

Function Existe_Relac(pMegaProcesso, pProcesso, pSubProcesso, pModulo, pAtividadeCarga, pAtual)
   str_SQL_Existe_Relac = ""
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " Select *  "
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " WHERE "
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & pMegaProcesso
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = " & pProcesso
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & pSubProcesso
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MODU_CD_MODULO = " & pModulo 
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = " & pAtividadeCarga
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = '" & pAtual & "'"
   
   'response.write str_SQL_Existe_Relac
   Set rdsExiste = conn_db.Execute(str_SQL_Existe_Relac)
   if rdsExiste.EOF then
      Existe_Relac = False
   else	  
      Existe_Relac = True
   end if
end function

Sub Grava_Delecao_Transa(strMP, strP, strSP, strM, strA, strT)
	str_SQL = ""
	str_SQL = str_SQL & " Delete from " & Session("PREFIXO") & "RELACAO_FINAL "
    str_SQL = str_SQL & " WHERE "
    str_SQL = str_SQL & " " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strMP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & strP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & strSP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MODU_CD_MODULO = " & strM 
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & strA
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & strT & "'"
    
    on error resume next
    
    Set rdsDeleta = conn_db.Execute(str_SQL)
    
    if err.number <> 0 then
    	tem_erro = tem_erro + 1
    	err.clear
    end if
    
	'response.Write(str_SQL)
	
	int_Tot_Trans_Exc = int_Tot_Trans_Exc + 1	

end sub

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
      - 
      <%'=str_MegaProcesso%>
      - 
      <%'=str_Processo%>
      - 
      <%'=str_SubProcesso%>
      - 
      <%'=str_AtividadeCarga%>
    </td>
    <td width="6%">&nbsp;</td>
    <td width="62%"> 
      <%'=str_AtividadeCarga%>
      - 
      <%'=str_Modulo%>
      - 
      <%'=str_Transacao%>
      - 
      <%'=i%>
      - 
      <%'=str_NovaTransacao%>
      - 
      <%'=str_Trata%>
      - 
      <%'=int_Total_Atividade%>
    </td>
    <td width="12%"><%=str_SQL_Nova_Ativ_Tran%></td>
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
    <td width="12%">&nbsp; </td>
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
    <td width="12%">&nbsp; </td>
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
    <td width="12%">&nbsp; </td>
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
    <td width="57%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de transa&ccedil;&otilde;es gravadas:<b><%=int_Total_AtividadeTrans%> </b></font></td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de transa&ccedil;&otilde;es exclu&iacute;das:</font><%=int_Tot_Trans_Exc%></td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%"> 
	<% If ls_Indice > 0 then %>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">T</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">ransa&ccedil;&otilde;es 
            com outros Megas:</font></td>
        </tr>
        <tr> 
          <td width="52%"> 
            <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">T</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">ransa&ccedil;&otilde;es</font></div>
          </td>
          <td width="48%"> 
            <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Mega</font></div>
          </td>
        </tr>
        <% ls_Loop = 0
		Do While ls_Indice > ls_Loop %>
        <tr> 
          <td width="52%"> 
            <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=ls_Trans_Exist(ls_Loop)%></font></b></div>
          </td>
          <td width="48%"> 
            <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=ls_Mega_Exist(ls_Loop)%></font></b></div>
          </td>
        </tr>
        <% ls_Loop = ls_Loop + 1
		Loop %>
      </table>
	  <% end if %>
    </td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%">&nbsp;</td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%"> 
      <table width="98%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="12%"><a href="selec_Mega_Proc_Sub_Processo.asp?txtOpc=2&selMegaProcesso=<%=str_MegaProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Mega Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="selec_Mega_Proc_Sub_Processo.asp?txtOpc=3&selProcesso=<%=str_MegaProcesso%>/<%=str_Processo%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="selec_Mega_Proc_Sub_Processo.asp?txtOpc=2&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona nova Atividade" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Selecuina 
            novo Sub Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans3.asp?txtOpc=3&selModulo=0&selAtividadeCarga=0&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo M&oacute;dulo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo </font><font color="#003300" face="Arial, Helvetica, sans-serif" size="2">Agrupamento 
            das Atividades</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans3.asp?txtOpc=3&selModulo=<%=str_Modulo%>&selAtividadeCarga=0&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            nova Atividade</font></td>
        </tr>
        <tr> 
          <td width="12%">&nbsp;</td>
          <td width="88%">&nbsp;</td>
        </tr>
      </table>
    </td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%">
    <%
    if tem_erro>0 then
    %>
    <p align="center"><b><font color="#FF0000" size="2" face="Verdana">Algumas Transações não puderam ser excluídas da Decomposição, por possuírem relação com outras entidades (Cenário, Função, ...). Por favor, verifique as relações e tente novamente.</font></b>
    <%
    end if
    %>
    </td>
    <td width="24%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->

<%
tem_erro = 0 

Dim str_Opc
dim str_MegaProcesso
dim str_Processo
dim str_SubProcesso
dim str_AtividadeCarga
dim str_Modulo
dim str_Transacao

dim int_Total_Atividade
dim int_Total_AtividadeTrans
Dim int_Tot_Trans_Exc
int_Tot_Trans_Exc = 0

str_Opc = Request("txtOpc")

str_MegaProcesso= Request("txtMP")
str_Processo = Request("txtP")
str_SubProcesso = Request("txtSP")
str_AtividadeCarga = Request("selAtividadeCarga")
str_Modulo = Request("selModulo")

str_DescMegaProcesso= Request("txtDsMP")
str_DescProcesso = Request("txtDsP")
str_DescSubProcesso = Request("txtDsSP")
str_DescAtividadeCarga = Request("txtDsA")
str_DescModulo = Request("txtDsM")

str_Transacao = Request("txtTranSelecionada")
str_NaoTransacao = Request("txtTranNaoSelecionada")
'response.Write "  Selec   "
'response.Write(str_Transacao)
'response.Write "  Nao Selec   "
'response.Write(str_NaoTransacao)

str_DsMP = Request("txtDsMP")

int_Total_Atividade = 0
int_Total_AtividadeTrans = 0

dim ls_Trans_Exist(50)
dim ls_Mega_Exist(50)
ls_Indice = 0
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

ls_Sequencia = 0

Sub Grava_Nova_Atividade_Trans(strMP, strP, strSP, strM, strA, strT)
    
	'str_SQL = ""
	'str_SQL = str_SQL & " SELECT TRAN_CD_TRANSACAO, MEPR_TX_DESC_MEGA_PROCESSO "
	'str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "RELACAO_FINAL, " & Session("PREFIXO") & "MEGA_PROCESSO "
	'str_SQL = str_SQL & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & strT & "'"
	'str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO <> " & strMP
	'str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
	''RESPONSE.WRITE str_SQL
	'Set rdsExiste = conn_db.Execute(str_SQL)
	'if not rdsExiste.EOF then
	'   ls_Trans_Exist(ls_Indice) = rdsExiste("TRAN_CD_TRANSACAO")
	'   ls_Mega_Exist(ls_Indice) = rdsExiste("MEPR_TX_DESC_MEGA_PROCESSO")
	'   ls_Indice = ls_Indice + 1
	'else
       ls_Sequencia = ls_Sequencia + 1
       str_SQL_Nova_Ativ_Tran = ""
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " INSERT INTO " & Session("PREFIXO") & "RELACAO_FINAL ( "
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " MEPR_CD_MEGA_PROCESSO "
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,PROC_CD_PROCESSO "
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,SUPR_CD_SUB_PROCESSO "
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,MODU_CD_MODULO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATCA_CD_ATIVIDADE_CARGA "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,TRAN_CD_TRANSACAO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,RELA_NR_SEQUENCIA "	
       str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_TX_OPERACAO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_CD_NR_USUARIO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ,ATUA_DT_ATUALIZACAO "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & " ) Values( "
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strMP & "," & strP & ","
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strSP & "," & strM & "," & strA & ","
	   str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & "'" & strT & "'," & ls_Sequencia  & ", 'I', '" & Session("CdUsuario") & "', GETDATE())" 
	   Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ_Tran)
	   strChave = CStr(strMP) & " " & CStr(strP) & " " & CStr(strSP) & " " & CStr(strM) & " " & CStr(strA) & CStr(strT)
	   'call grava_log(strChave,"RELACAO_FINAL","I",0)
       int_Total_AtividadeTrans = int_Total_AtividadeTrans + 1
	'end if
end sub

Function Existe_Atividade_Trans(pMegaProcesso, pProcesso, pSubProcesso, pModulo, pAtividadeCarga, pAtual)
   str_SQL_Deleta_Ativ_Tran = ""
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " Select *  "
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " FROM " & Session("PREFIXO") & "RELACAO_FINAL "
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " WHERE "
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & pMegaProcesso
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & pProcesso
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & pSubProcesso
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MODU_CD_MODULO = " & pModulo 
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & pAtividadeCarga
   str_SQL_Deleta_Ativ_Tran = str_SQL_Deleta_Ativ_Tran & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & pAtual & "'"
   'response.write str_SQL_Deleta_Ativ_Tran
   Set rdsExiste = conn_db.Execute(str_SQL_Deleta_Ativ_Tran)
   if rdsExiste.EOF then
      Existe_Atividade_Trans = False
   else	  
      Existe_Atividade_Trans = True
   end if
end function

'guarda o conteúdo da String
str_valor = str_Transacao

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
if tamanho > 0 then
'início da Rotina
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        if Existe_Atividade_Trans(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual) = false then
           'Aqui entra o que vc quer fazer com o caracter em questão!
	       call Grava_Nova_Atividade_Trans(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual)
        end if
        quantos = 0
    End If
    contador = contador + 1
Loop
end if


'guarda o conteúdo da String
str_valor = str_NaoTransacao
'response.Write(str_valor)
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
if tamanho > 0 then
'início da Rotina
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        if Existe_Relac(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual) = False then
           If Existe_em_Rel_Final(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual) = True then
              'Aqui entra o que vc quer fazer com o caracter em questão!
	          call Grava_Delecao_Transa(str_MegaProcesso, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual)
           end if
		end if
        quantos = 0
    End If
    contador = contador + 1
Loop
end if

Function Existe_em_Rel_Final(strMP, strP, strSP, strM, strA, strT)
	str_SQL = ""
	str_SQL = str_SQL & " Select *  from " & Session("PREFIXO") & "RELACAO_FINAL "
    str_SQL = str_SQL & " WHERE "
    str_SQL = str_SQL & " " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strMP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & strP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & strSP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MODU_CD_MODULO = " & strM 
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & strA
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & strT & "'"
    Set rsExiste_em_Rel_Final = conn_db.Execute(str_SQL)
	'response.Write(str_SQL)
    'response.write str_SQL
    Set rsExiste_em_Rel_Final = conn_db.Execute(str_SQL)
    if rsExiste_em_Rel_Final.EOF then
       Existe_em_Rel_Final = False
    else	  
       Existe_em_Rel_Final = True
    end if
end function

Function Existe_Relac(pMegaProcesso, pProcesso, pSubProcesso, pModulo, pAtividadeCarga, pAtual)
   str_SQL_Existe_Relac = ""
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " Select *  "
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " WHERE "
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & pMegaProcesso
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = " & pProcesso
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & pSubProcesso
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MODU_CD_MODULO = " & pModulo 
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = " & pAtividadeCarga
   str_SQL_Existe_Relac = str_SQL_Existe_Relac & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = '" & pAtual & "'"
   
   'response.write str_SQL_Existe_Relac
   Set rdsExiste = conn_db.Execute(str_SQL_Existe_Relac)
   if rdsExiste.EOF then
      Existe_Relac = False
   else	  
      Existe_Relac = True
   end if
end function

Sub Grava_Delecao_Transa(strMP, strP, strSP, strM, strA, strT)
	str_SQL = ""
	str_SQL = str_SQL & " Delete from " & Session("PREFIXO") & "RELACAO_FINAL "
    str_SQL = str_SQL & " WHERE "
    str_SQL = str_SQL & " " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strMP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & strP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & strSP
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.MODU_CD_MODULO = " & strM 
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & strA
    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & strT & "'"
    
    on error resume next
    
    Set rdsDeleta = conn_db.Execute(str_SQL)
    
    if err.number <> 0 then
    	tem_erro = tem_erro + 1
    	err.clear
    end if
    
	'response.Write(str_SQL)
	
	int_Tot_Trans_Exc = int_Tot_Trans_Exc + 1	

end sub

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
      - 
      <%'=str_MegaProcesso%>
      - 
      <%'=str_Processo%>
      - 
      <%'=str_SubProcesso%>
      - 
      <%'=str_AtividadeCarga%>
    </td>
    <td width="6%">&nbsp;</td>
    <td width="62%"> 
      <%'=str_AtividadeCarga%>
      - 
      <%'=str_Modulo%>
      - 
      <%'=str_Transacao%>
      - 
      <%'=i%>
      - 
      <%'=str_NovaTransacao%>
      - 
      <%'=str_Trata%>
      - 
      <%'=int_Total_Atividade%>
    </td>
    <td width="12%"><%=str_SQL_Nova_Ativ_Tran%></td>
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
    <td width="12%">&nbsp; </td>
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
    <td width="12%">&nbsp; </td>
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
    <td width="12%">&nbsp; </td>
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
    <td width="57%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de transa&ccedil;&otilde;es gravadas:<b><%=int_Total_AtividadeTrans%> </b></font></td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de transa&ccedil;&otilde;es exclu&iacute;das:</font><%=int_Tot_Trans_Exc%></td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%"> 
	<% If ls_Indice > 0 then %>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">T</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">ransa&ccedil;&otilde;es 
            com outros Megas:</font></td>
        </tr>
        <tr> 
          <td width="52%"> 
            <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">T</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">ransa&ccedil;&otilde;es</font></div>
          </td>
          <td width="48%"> 
            <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Mega</font></div>
          </td>
        </tr>
        <% ls_Loop = 0
		Do While ls_Indice > ls_Loop %>
        <tr> 
          <td width="52%"> 
            <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=ls_Trans_Exist(ls_Loop)%></font></b></div>
          </td>
          <td width="48%"> 
            <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=ls_Mega_Exist(ls_Loop)%></font></b></div>
          </td>
        </tr>
        <% ls_Loop = ls_Loop + 1
		Loop %>
      </table>
	  <% end if %>
    </td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%">&nbsp;</td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%"> 
      <table width="98%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="12%"><a href="selec_Mega_Proc_Sub_Processo.asp?txtOpc=2&selMegaProcesso=<%=str_MegaProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Mega Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="selec_Mega_Proc_Sub_Processo.asp?txtOpc=3&selProcesso=<%=str_MegaProcesso%>/<%=str_Processo%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="selec_Mega_Proc_Sub_Processo.asp?txtOpc=2&selAtividadeCarga=<%=str_AtividadeCarga%>&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona nova Atividade" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Selecuina 
            novo Sub Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans3.asp?txtOpc=3&selModulo=0&selAtividadeCarga=0&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo M&oacute;dulo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo </font><font color="#003300" face="Arial, Helvetica, sans-serif" size="2">Agrupamento 
            das Atividades</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="form_relaciona_ativ_trans3.asp?txtOpc=3&selModulo=<%=str_Modulo%>&selAtividadeCarga=0&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            nova Atividade</font></td>
        </tr>
        <tr> 
          <td width="12%">&nbsp;</td>
          <td width="88%">&nbsp;</td>
        </tr>
      </table>
    </td>
    <td width="24%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="57%">
    <%
    if tem_erro>0 then
    %>
    <p align="center"><b><font color="#FF0000" size="2" face="Verdana">Algumas Transações não puderam ser excluídas da Decomposição, por possuírem relação com outras entidades (Cenário, Função, ...). Por favor, verifique as relações e tente novamente.</font></b>
    <%
    end if
    %>
    </td>
    <td width="24%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
</html>