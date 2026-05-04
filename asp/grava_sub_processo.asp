<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
dim int_Total_Tot_Empresa
dim str_MegaProcesso
dim str_Processo
dim str_SubProceso
dim str_SQL_Nova_Sub_Empr
dim int_MaxSubProcesso

int_Total_Tot_Empresa = 0

set Conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Opc = Request("txtOpc")
str_MegaProcesso = request("selMegaProcesso")
str_Processo = request("selProcesso")
str_Empresa_Unid = Request("txtEmpSelecionada")
str_Empresa_Unid_Original = str_Empresa_Unid
if Len(str_Empresa_Unid) > 2 then    
   if InStr(str_Empresa_Unid,",10") <> 0 or InStr(str_Empresa_Unid,",11") <> 0 or InStr(str_Empresa_Unid,",12") <> 0 then      
      str_Empresa_Unid = str_Empresa_Unid + ","
	  if InStr(str_Empresa_Unid,",1,") <> 0 then
         msg = "tem e achou o 1"
	  else
	     msg = "tem e não achou o 1"
		 str_Empresa_Unid = str_Empresa_Unid_Original & ",1" 
	  end if
   else
      msg = "não tem 10/11/12"
   end if	  
end if
str_NovoSubProceso = request("txtNovoSubProc1")
str_Seq = request("txtSeq1")
if str_Seq = "" then
   str_Seq = "0"
end if
Sub Grava_Novo_Sub_Processo(str_NovoSubProcesso, ls_Seq)
	str_SQL_Sub_Proc = ""
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " SELECT "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " MAX(SUPR_CD_SUB_PROCESSO) AS MAXIMO_SUB "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " FROM " & Session("PREFIXO") & "SUB_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO, "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " PROC_CD_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND PROC_CD_PROCESSO = " & str_Processo
	Set rdsMaxSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
    if rdsMaxSubProcesso.EOF then
	   int_MaxSubProcesso = 1	
    else
   	   int_MaxSubProcesso = rdsMaxSubProcesso("MAXIMO_SUB") + 1	
    end if
	rdsMaxSubProcesso.Close
	set rdsMaxSubProcesso = Nothing
    
    impacto = cInt(request("valImpacto"))
    
    str_SQL_Sub_Proc = ""
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " INSERT INTO " & Session("PREFIXO") & "SUB_PROCESSO ( "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,PROC_CD_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_CD_SUB_PROCESSO "
		    
	if impacto<>0 then
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_TX_IMPACTO "
	end if
	
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_TX_DESC_SUB_PROCESSO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_NR_SEQUENCIA "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_TX_OPERACAO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ) Values( "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & str_MegaProcesso & "," & str_Processo & "," & int_MaxSubProcesso & ","
	
	if impacto<>0 then
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & "" & impacto & ",'" & Ucase(str_NovoSubProcesso) & "'," & ls_Seq & ", 'I', 'XXXX', GETDATE())" 
	ELSE
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & "'" & Ucase(str_NovoSubProcesso) & "'," & ls_Seq & ", 'I', 'XXXX', GETDATE())" 
	END IF
	
	Set rdsNovoSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
	
	strChave = CStr(str_MegaProcesso) & " " & CStr(str_Processo) & " " & CStr(int_MaxSubProcesso) 
	'call grava_log(strChave,"SUB_PROCESSO","I",0)

end sub

Sub Grava_Novo_Sub_Processo_Empresa(strMP, strP, strSP, strEU)
    str_SQL_Nova_Sub_Empr = ""
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " INSERT INTO " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE ( "
    str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,PROC_CD_PROCESSO "
    str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,SUPR_CD_SUB_PROCESSO "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,EMPR_CD_NR_EMPRESA "
    str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,ATUA_TX_OPERACAO "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ) Values( "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & strMP & "," & strP & ","
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & strSP & "," & strEU & ","
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " 'I', 'XXXX', GETDATE())" 
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Sub_Empr)
    int_Total_Tot_Empresa = int_Total_Tot_Empresa + 1
end sub

if str_NovoSubProceso <> "" then
   call Grava_Novo_Sub_Processo(Trim(str_NovoSubProceso),str_Seq)
end if


'guarda o conteúdo da String
str_valor = str_Empresa_Unid

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
        
	' Aqui entra o que vc quer fazer com o caracter em questão!
	   call Grava_Novo_Sub_Processo_Empresa(str_MegaProcesso, str_Processo, int_MaxSubProcesso, str_atual)

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
<script language="JavaScript">
<!--
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frm1.selMegaProcesso.value+"'");
}
function Confirma2() 
{ 
	  document.frm1.submit();
}
function Confirma() 
{ 
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
	 else
     {
	  document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top" height="65"> 
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
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
    <td width="70%"><%'=str_Empresa_Unid%>
    </td>
    <td width="16%"><%'=msg%></td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de registro gravado:</font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=int_Total_Tot_Empresa%></font></td>
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
      <%'=str_Opc%>
    </td>
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
      <%if str_Opc <> "1" then %>
      <table width="82%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="JavaScript:history.go(-2)"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Sele&ccedil;&atilde;o de Sub-Processo </font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
      <p> 
        <% end if %>
      </p>
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      <%if str_Opc = "1"  then %>
      <table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="cadas_sub_processo.asp?txtOpc=1&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Novo Sub-Processo </font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
      <% end if %>
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
dim int_Total_Tot_Empresa
dim str_MegaProcesso
dim str_Processo
dim str_SubProceso
dim str_SQL_Nova_Sub_Empr
dim int_MaxSubProcesso

int_Total_Tot_Empresa = 0

set Conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Opc = Request("txtOpc")
str_MegaProcesso = request("selMegaProcesso")
str_Processo = request("selProcesso")
str_Empresa_Unid = Request("txtEmpSelecionada")
str_Empresa_Unid_Original = str_Empresa_Unid
if Len(str_Empresa_Unid) > 2 then    
   if InStr(str_Empresa_Unid,",10") <> 0 or InStr(str_Empresa_Unid,",11") <> 0 or InStr(str_Empresa_Unid,",12") <> 0 then      
      str_Empresa_Unid = str_Empresa_Unid + ","
	  if InStr(str_Empresa_Unid,",1,") <> 0 then
         msg = "tem e achou o 1"
	  else
	     msg = "tem e não achou o 1"
		 str_Empresa_Unid = str_Empresa_Unid_Original & ",1" 
	  end if
   else
      msg = "não tem 10/11/12"
   end if	  
end if
str_NovoSubProceso = request("txtNovoSubProc1")
str_Seq = request("txtSeq1")
if str_Seq = "" then
   str_Seq = "0"
end if
Sub Grava_Novo_Sub_Processo(str_NovoSubProcesso, ls_Seq)
	str_SQL_Sub_Proc = ""
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " SELECT "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " MAX(SUPR_CD_SUB_PROCESSO) AS MAXIMO_SUB "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " FROM " & Session("PREFIXO") & "SUB_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO, "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " PROC_CD_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND PROC_CD_PROCESSO = " & str_Processo
	Set rdsMaxSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
    if rdsMaxSubProcesso.EOF then
	   int_MaxSubProcesso = 1	
    else
   	   int_MaxSubProcesso = rdsMaxSubProcesso("MAXIMO_SUB") + 1	
    end if
	rdsMaxSubProcesso.Close
	set rdsMaxSubProcesso = Nothing
    
    impacto = cInt(request("valImpacto"))
    
    str_SQL_Sub_Proc = ""
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " INSERT INTO " & Session("PREFIXO") & "SUB_PROCESSO ( "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,PROC_CD_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_CD_SUB_PROCESSO "
		    
	if impacto<>0 then
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_TX_IMPACTO "
	end if
	
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_TX_DESC_SUB_PROCESSO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_NR_SEQUENCIA "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_TX_OPERACAO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ) Values( "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & str_MegaProcesso & "," & str_Processo & "," & int_MaxSubProcesso & ","
	
	if impacto<>0 then
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & "" & impacto & ",'" & Ucase(str_NovoSubProcesso) & "'," & ls_Seq & ", 'I', 'XXXX', GETDATE())" 
	ELSE
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & "'" & Ucase(str_NovoSubProcesso) & "'," & ls_Seq & ", 'I', 'XXXX', GETDATE())" 
	END IF
	
	Set rdsNovoSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
	
	strChave = CStr(str_MegaProcesso) & " " & CStr(str_Processo) & " " & CStr(int_MaxSubProcesso) 
	'call grava_log(strChave,"SUB_PROCESSO","I",0)

end sub

Sub Grava_Novo_Sub_Processo_Empresa(strMP, strP, strSP, strEU)
    str_SQL_Nova_Sub_Empr = ""
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " INSERT INTO " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE ( "
    str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,PROC_CD_PROCESSO "
    str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,SUPR_CD_SUB_PROCESSO "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,EMPR_CD_NR_EMPRESA "
    str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,ATUA_TX_OPERACAO "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ) Values( "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & strMP & "," & strP & ","
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & strSP & "," & strEU & ","
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " 'I', 'XXXX', GETDATE())" 
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Sub_Empr)
    int_Total_Tot_Empresa = int_Total_Tot_Empresa + 1
end sub

if str_NovoSubProceso <> "" then
   call Grava_Novo_Sub_Processo(Trim(str_NovoSubProceso),str_Seq)
end if


'guarda o conteúdo da String
str_valor = str_Empresa_Unid

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
        
	' Aqui entra o que vc quer fazer com o caracter em questão!
	   call Grava_Novo_Sub_Processo_Empresa(str_MegaProcesso, str_Processo, int_MaxSubProcesso, str_atual)

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
<script language="JavaScript">
<!--
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frm1.selMegaProcesso.value+"'");
}
function Confirma2() 
{ 
	  document.frm1.submit();
}
function Confirma() 
{ 
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
	 else
     {
	  document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top" height="65"> 
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
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
    <td width="70%"><%'=str_Empresa_Unid%>
    </td>
    <td width="16%"><%'=msg%></td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de registro gravado:</font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=int_Total_Tot_Empresa%></font></td>
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
      <%'=str_Opc%>
    </td>
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
      <%if str_Opc <> "1" then %>
      <table width="82%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="JavaScript:history.go(-2)"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Sele&ccedil;&atilde;o de Sub-Processo </font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
      <p> 
        <% end if %>
      </p>
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      <%if str_Opc = "1"  then %>
      <table width="80%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="cadas_sub_processo.asp?txtOpc=1&selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Novo Sub-Processo </font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
      <% end if %>
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
