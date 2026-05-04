<%@LANGUAGE="VBSCRIPT"%> 
 
<%
dim int_Total_Atividade
dim int_Total_AtividadeTrans

if (Request("txtMegaProcesso") <> "") then 
    str_MegaProcesso = Request("txtMegaProcesso")
else
    str_MegaProcesso = "não passado"
end if

if (Request("txtDescMegaProcesso") <> "") then 
    str_DescMegaProcesso = Request("txtDescMegaProcesso")
else
    str_DescMegaProcesso = "não passado"
end if

if (Request("txtFuncao") <> "") then 
    str_Funcao = Request("txtFuncao")
else
    str_Funcao = "não passado"
end if

if (Request("txtDescFuncao") <> "") then 
    str_DescFuncao = Request("txtDescFuncao")
else
    str_DescFuncao = "não passado"
end if

if (Request("selMegaProcesso2") <> "") then 
    str_MegaProcesso2 = Request("selMegaProcesso2")
else
    str_MegaProcesso2 = "não passado"
end if

if (Request("txtDescMegaProcesso2") <> "") then 
    str_DescMegaProcesso2 = Request("txtDescMegaProcesso2")
else
    str_DescMegaProcesso2 = "não passado"
end if

if (Request("selProcesso") <> "") then 
    str_Processo = Request("selProcesso")
else
    str_Processo = "não passado"
end if

if (Request("txtDescProcesso") <> "") then 
    str_DescProcesso = Request("txtDescProcesso")
else
    str_DescProcesso = "não passado"
end if

if (Request("selSubProcesso") <> "") then 
    str_SubProcesso = Request("selSubProcesso")
else
    str_SubProcesso = "0"
end if

if (Request("txtDescSubProcesso") <> "") then 
    str_DescSubProcesso = Request("txtDescSubProcesso")
else
    str_DescSubProcesso = "0"
end if

if (Request("selAtividadeCarga") <> "") then 
    str_AtividadeCarga = Request("selAtividadeCarga")
else
    str_AtividadeCarga = "não passado"
end if

if (Request("txtDescAtividadeCarga") <> "") then 
    str_DescAtividadeCarga = Request("txtDescAtividadeCarga")
else
    str_DescAtividadeCarga = "não passado"
end if

if (Request("txtTranSelecionada") <> "") then 
    str_TranSelecionada = Request("txtTranSelecionada")
else
    str_TranSelecionada = "não passado"
end if

int_Total_Transacoes_Exc = 0
int_Total_Transacoes_Cad = 0

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

Sub Grava_Nova_Funcao_Trans(strFn, strMP, strP, strSP, strM, strA, strT)
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
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & "'" & strFN & "',"
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strMP & "," & strP & ","
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & strSP & "," & 1 & "," & strA & ","
	str_SQL_Nova_Ativ_Tran = str_SQL_Nova_Ativ_Tran & "'" & strT & "'," & "'I', '" & Session("CdUsuario") & "', GETDATE())" 
	
	'response.write str_SQL_Nova_Ativ_Tran
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Ativ_Tran)
	'call grava_log(strFN,"" & Session("PREFIXO") & "FUN_NEG_TRANSACAO","I",0)
    int_Total_Transacoes_Cad = int_Total_Transacoes_Cad + 1
	
end sub

Sub Grava_Funcao_Trans_Relacionadas(strFn2, strMP2, strSP2, strA2, strT2)
    str_SQL_Nova_Trans_Relacionadas = ""
	str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " Select "
    str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO "
    str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " ," & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO "
    str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " ," & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA "		
    str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " ," & Session("PREFIXO") & "RELACAO_FINAL.MODU_CD_MODULO "
	str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " ," & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO "	
	str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " FROM " & Session("PREFIXO") & "RELACAO_FINAL "
    str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " WHERE "			
	str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & strMP2
	str_SQL_Nova_Trans_Relacionadas = str_SQL_Nova_Trans_Relacionadas & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = '" & strT2 & "'"
		
    'response.write str_SQL_Nova_Trans_Relacionadas

	Set rdsTrans_Relacionadas = conn_db.Execute(str_SQL_Nova_Trans_Relacionadas)

	
	Do while not rdsTrans_Relacionadas.EOF
	    str_SQL = ""
		str_SQL = str_SQL & " Select MEPR_CD_MEGA_PROCESSO "
		str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
	    str_SQL = str_SQL & " WHERE "
	    str_SQL = str_SQL & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '" & strFn2 & "'"
	    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & strMP2
	    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = " & rdsTrans_Relacionadas("PROC_CD_PROCESSO")
	    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & rdsTrans_Relacionadas("SUPR_CD_SUB_PROCESSO")
	    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = " & rdsTrans_Relacionadas("ATCA_CD_ATIVIDADE_CARGA")
	    str_SQL = str_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = '" & strT2 & "'"
		'response.write str_SQL
		Set rdsExiste = conn_db.Execute(str_SQL)
		if rdsExiste.EOF then
		   call Grava_Nova_Funcao_Trans(strFn2, strMP2, rdsTrans_Relacionadas("PROC_CD_PROCESSO"), rdsTrans_Relacionadas("SUPR_CD_SUB_PROCESSO"), rdsTrans_Relacionadas("MODU_CD_MODULO"), rdsTrans_Relacionadas("ATCA_CD_ATIVIDADE_CARGA"), strT2)
		end if   
		rdsExiste.close
    	rdsTrans_Relacionadas.movenext
	Loop
	rdsTrans_Relacionadas.close
	set rdsTrans_Relacionadas = Nothing
			
end sub

Sub Deleta_Funcao_Transacao(strFn, strMP, strP, strSP, strA, strT)
	str_SQL_Deleta_Func_Tran = ""
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " DELETE "
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO "
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " WHERE "
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '" & strFn & "'"
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & strMP
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = " & strP
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & strSP
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = " & strA
	str_SQL_Deleta_Func_Tran = str_SQL_Deleta_Func_Tran & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = '" & strT & "'"

	Set rdsDeleta = conn_db.Execute(str_SQL_Deleta_Func_Tran)
	'call grava_log(strFN,"" & Session("PREFIXO") & "FUN_NEG_TRANSACAO","D",0)

	int_Total_Transacoes_Exc = int_Total_Transacoes_Exc + 1

End Sub

'************  CASO SEJA UMA ALTERAÇÃO - EXCLUI AS TRANSAÇÃOES QUE NÃO FORAM SELECIONADAS ****************
str_SQL_Transacao_Cad = ""
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " SELECT  DISTINCT  "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " TRAN_CD_TRANSACAO "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso2
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND PROC_CD_PROCESSO = " & str_Processo
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'"

Set rdsTransacao_Cad = Conn_db.Execute(str_SQL_Transacao_Cad)
str_TranCadastradas = ","
ls_int_Total_Deletada = 0 
do while Not rdsTransacao_Cad.EOF 
   ls_str_Cd_Transacao = "," & rdsTransacao_Cad("TRAN_CD_TRANSACAO") & ","
   str_TranCadastradas = str_TranCadastradas & rdsTransacao_Cad("TRAN_CD_TRANSACAO") & ","
   if InStr(str_TranSelecionada & ",",ls_str_Cd_Transacao) = 0 then
      Call Deleta_Funcao_Transacao(str_Funcao, str_MegaProcesso2, str_Processo, str_SubProcesso, str_AtividadeCarga, rdsTransacao_Cad("TRAN_CD_TRANSACAO")) 	  
	  ls_int_Total_Deletada = ls_int_Total_Deletada + 1
   end if
   rdsTransacao_Cad.movenext
Loop

rdsTransacao_Cad.close
set rdsTransacao_Cad = Nothing
'***********************  FIM EXCLUSÃO ****************************

'guarda o conteúdo da String
str_valor = str_TranSelecionada

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
        ' **** verifica nas transações cadastradas se não existe aquela selecionada 
		if InStr(str_TranCadastradas,"," & str_atual & ",") = 0 then
          call Grava_Nova_Funcao_Trans(str_Funcao, str_MegaProcesso2, str_Processo, str_SubProcesso, str_Modulo, str_AtividadeCarga, str_atual)
		   call Grava_Funcao_Trans_Relacionadas(str_Funcao, str_MegaProcesso2, str_SubProcesso, str_AtividadeCarga, str_atual)
        end if          
        quantos = 0
    End If
    contador = contador + 1
Loop
end if
aa = "," & str_atual & ","
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
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--
function MM_goToURL1() { //v3.0
   var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+txtMegaProcesso.value+"'");
}
function MM_goToURL2() { //v3.0
   var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+txtMegaProcesso.value+"&selFuncao="+txtFuncao.value+"'");
}
function MM_goToURL3() { //v3.0
   var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+txtMegaProcesso.value+"&selFuncao="+txtFuncao.value+"&selMegaProcesso2="+txtMegaProcesso2.value+"'");
}
function MM_goToURL4() { //v3.0
   var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+txtMegaProcesso.value+"&selFuncao="+txtFuncao.value+"&selMegaProcesso2="+txtMegaProcesso2.value+"&selProcesso="+txtProcesso.value+"'");
}
function MM_goToURL5() { //v3.0
   var i, args=MM_goToURL5.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+txtMegaProcesso.value+"&selFuncao="+txtFuncao.value+"&selMegaProcesso2="+txtMegaProcesso2.value+"&selProcesso="+txtProcesso.value+"&selSubProcesso="+txtSubProcesso.value+"'");
}

//-->
</script>

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
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a> 
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="91%" border="0" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td width="19%">&nbsp;</td>
    <td colspan="2"><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Grava 
      relacionamento de Fun&ccedil;&atilde;o R/3 e transa&ccedil;&atilde;o</font></td>
    <td width="2%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%"> 
      <%'=str_Opc%>
      <%'=str_MegaProcesso%>
      <%'=str_Processo%>
      <%'=str_SubProcesso%>
      <%'=str_AtividadeCarga%>
    </td>
    <td width="6%">&nbsp;</td>
    <td width="66%"> 
      <%'=str_AtividadeCarga%>
      <%'=str_Modulo%>
      <%'=str_Transacao%>
      <%'=i%>
      <%'=str_NovaTransacao%>
      <%'=str_Trata%>
      <%'=int_Total_Atividade%>
    </td>
    <td width="2%"> 
      <%'=str_TranSelecionada%>
    </td>
  </tr>
  <tr> 
    <td width="19%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Mega 
        Processo: </font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_MegaProcesso%></font> </font></div>
    </td>
    <td width="66%"><b><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      &nbsp;&nbsp;- <%=str_DescMegaProcesso%> 
      <input type="hidden" name="txtMegaProcesso" value="<%=str_MegaProcesso%>">
      </font></font></b></td>
    <td width="2%"> 
      <%'=ls_int_Total_Deletada%>
    </td>
  </tr>
  <tr> 
    <td width="19%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Fun&ccedil;&atilde;o R/3</font><font face="Arial, Helvetica, sans-serif" size="2">: 
        </font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_Funcao%></font> </font></div>
    </td>
    <td width="66%"><b><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp; 
      - </font><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_DescFuncao%></font><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="hidden" name="txtFuncao" value="<%=str_Funcao%>">
      </font></font></font></font></font></font></b></td>
    <td width="2%"> 
      <%'=str_TranCadastradas%>
    </td>
  </tr>
  <tr bgcolor="#0099CC"> 
    <td width="19%" height="4"></td>
    <td width="6%" height="4"></td>
    <td width="66%" height="4"></td>
    <td width="2%" height="4"></td>
  </tr>
  <tr> 
    <td width="19%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Mega 
        Processo: </font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_MegaProcesso2%></font> </font></div>
    </td>
    <td width="66%"><b><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      &nbsp;&nbsp;- <%=str_DescMegaProcesso%></font><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="hidden" name="txtMegaProcesso2" value="<%=str_MegaProcesso2%>">
      </font></font></font></b></td>
    <td width="2%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Processo: 
        </font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_Processo%></font> </font></div>
    </td>
    <td width="66%"><b><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescProcesso%></font><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="hidden" name="txtProcesso" value="<%=str_Processo%>">
      </font></font></font></b></td>
    <td width="2%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Sub 
        Processo: </font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_SubProcesso%></font> </font></div>
    </td>
    <td width="66%"><b><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescSubProcesso%></font><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="hidden" name="txtSubProcesso" value="<%=str_SubProcesso%>">
      </font></font></font></b></td>
    <td width="2%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade: 
        </font></font></div>
    </td>
    <td width="6%"> 
      <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=str_AtividadeCarga%></font> </font></div>
    </td>
    <td width="66%"><b><font color="#003366"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;</font></font><font face="Arial, Helvetica, sans-serif" size="2">- 
      <%=str_DescAtividadeCarga%></font><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"> 
      <input type="hidden" name="txtAtividadeCarga" value="<%=str_AtividadeCarga%>">
      </font></font></font></b></td>
    <td width="2%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%" height="4" bgcolor="#0099CC"></td>
    <td width="6%" height="4" bgcolor="#0099CC"></td>
    <td width="66%" height="4" bgcolor="#0099CC"></td>
    <td width="2%" height="4" bgcolor="#0099CC"></td>
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
    <td width="47%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de transa&ccedil;&otilde;es gravadas:<%=int_Total_Transacoes_Cad%> </font></td>
    <td width="34%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="19%">&nbsp;</td>
    <td width="47%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de transa&ccedil;&otilde;es exclu&iacute;das:<%=int_Total_Transacoes_Exc%> </font></td>
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
          <td width="12%"><a href="javascript:MM_goToURL1('self','selec_Mega_Funcao.asp',this)"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            nova Fun&ccedil;&atilde;o R/3</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="javascript:MM_goToURL2('self','cad_funcao_transacao.asp',this)"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Mega Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="javascript:MM_goToURL3('self','cad_funcao_transacao.asp',this)"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            novo Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="javascript:MM_goToURL4('self','cad_funcao_transacao.asp',this)"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Selecuina 
            novo Sub Processo</font></td>
        </tr>
        <tr> 
          <td width="12%"><a href="javascript:MM_goToURL5('self','cad_funcao_transacao.asp',this)"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a></td>
          <td width="88%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Seleciona 
            nova Atividade</font></td>
        </tr>
        <tr> 
          <td width="12%">&nbsp;</td>
          <td width="88%">&nbsp;</td>
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
</body>
</html>


