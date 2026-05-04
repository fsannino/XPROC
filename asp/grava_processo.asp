<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
dim str_SQL_Proc 
dim int_Total_Gravado
dim str_MegaProcesso

str_Opc = Request("txtOpc")
str_MegaProcesso = request("selMegaProcesso")


Application("MegaProcesso") = str_MegaProcesso


Sub Grava_Novo_Processo(str_NovoProcesso,ls_Seq)

	set Conn_db = Server.CreateObject("ADODB.Connection")
	Conn_db.Open Session("Conn_String_Cogest_Gravacao")

	str_SQL_Proc = ""
	str_SQL_Proc = str_SQL_Proc & " SELECT "
	str_SQL_Proc = str_SQL_Proc & " MAX(PROC_CD_PROCESSO) AS MAX_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " FROM " & Session("PREFIXO") & "PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
	
	Set rdsMaxProcesso = Conn_db.Execute(str_SQL_Proc)
	
	if rdsMaxProcesso.EOF then
	   int_MaxProcesso = 1	
	else
	   int_MaxProcesso = rdsMaxProcesso("MAX_PROCESSO") + 1	
	end if
	rdsMaxProcesso.Close
	set rdsMaxProcesso = Nothing
    str_SQL_Proc = ""
	str_SQL_Proc = str_SQL_Proc & " INSERT INTO " & Session("PREFIXO") & "PROCESSO ( "
    str_SQL_Proc = str_SQL_Proc & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,PROC_CD_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,PROC_TX_DESC_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " ,PROC_NR_SEQUENCIA "
    str_SQL_Proc = str_SQL_Proc & " ,ATUA_TX_OPERACAO "
	str_SQL_Proc = str_SQL_Proc & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Proc = str_SQL_Proc & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Proc = str_SQL_Proc & " ) Values( "
	str_SQL_Proc = str_SQL_Proc & str_MegaProcesso & "," & int_MaxProcesso & ","
	str_SQL_Proc = str_SQL_Proc & "'" & UCase(str_NovoProcesso) & "'," & ls_Seq & ", 'I', 'XXXX', GETDATE())" 
	Set rdsNovoProcesso = Conn_db.Execute(str_SQL_Proc)

    strChave = CStr(str_MegaProcesso) & CStr(int_MaxProcesso) ' & CStr(int_SubProcesso) & CStr(int_MaxAtividade) ' & CStr(strEU)
	'call grava_log(strChave,"PROCESSO","I",0)
		
    int_Total_Gravado = int_Total_Gravado + 1
	conn_db.Close
	set conn_db = Nothing
end sub

if request("txtNovoProc1") <> "" then
   str_Seq = request("txtSeq1")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc1"),str_Seq)
end if
if request("txtNovoProc2") <> "" then
   str_Seq = request("txtSeq2")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc2"),str_Seq)
end if
if request("txtNovoProc3") <> "" then
   str_Seq = request("txtSeq3")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc3"),str_Seq)
end if
if request("txtNovoProc4") <> "" then
   str_Seq = request("txtSeq4")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc4"),str_Seq)
end if
if request("txtNovoProc5") <> "" then
   str_Seq = request("txtSeq5")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc5"),str_Seq)
end if
if request("txtNovoProc6") <> "" then
   str_Seq = request("txtSeq6")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc6"),str_Seq)
end if
if request("txtNovoProc7") <> "" then
   str_Seq = request("txtSeq7")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc7"),str_Seq)
end if
if request("txtNovoProc8") <> "" then
   str_Seq = request("txtSeq8")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc8"),str_Seq)
end if
if request("txtNovoProc9") <> "" then
   str_Seq = request("txtSeq9")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc9"),str_Seq)
end if
if request("txtNovoProc10") <> "" then
   str_Seq = request("txtSeq10")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc10"),str_Seq)
end if

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
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
<table width="75%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%"><%'=Application("MegaProcesso")%></td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de registro gravado:</font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=int_Total_Gravado%></font></td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%"><%'=str_Opc%></td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">
        <%if str_Opc <> "1" then %>
      <table width="74%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="JavaScript:history.go(-2)"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Sele&ccedil;&atilde;o de Processo </font></td>
        </tr>
        <tr>
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
		<% end if %>
    </td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">
	 <%if str_Opc = "1" then %>
      <table width="75%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="cadas_processo.asp?txtOpc=1&selMegaProcesso=<%=str_MegaProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Novo Processo</font></td>
        </tr>
        <tr>
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
	  <% end if %>
    </td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
dim str_SQL_Proc 
dim int_Total_Gravado
dim str_MegaProcesso

str_Opc = Request("txtOpc")
str_MegaProcesso = request("selMegaProcesso")


Application("MegaProcesso") = str_MegaProcesso


Sub Grava_Novo_Processo(str_NovoProcesso,ls_Seq)

	set Conn_db = Server.CreateObject("ADODB.Connection")
	Conn_db.Open Session("Conn_String_Cogest_Gravacao")

	str_SQL_Proc = ""
	str_SQL_Proc = str_SQL_Proc & " SELECT "
	str_SQL_Proc = str_SQL_Proc & " MAX(PROC_CD_PROCESSO) AS MAX_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " FROM " & Session("PREFIXO") & "PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
	
	Set rdsMaxProcesso = Conn_db.Execute(str_SQL_Proc)
	
	if rdsMaxProcesso.EOF then
	   int_MaxProcesso = 1	
	else
	   int_MaxProcesso = rdsMaxProcesso("MAX_PROCESSO") + 1	
	end if
	rdsMaxProcesso.Close
	set rdsMaxProcesso = Nothing
    str_SQL_Proc = ""
	str_SQL_Proc = str_SQL_Proc & " INSERT INTO " & Session("PREFIXO") & "PROCESSO ( "
    str_SQL_Proc = str_SQL_Proc & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,PROC_CD_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,PROC_TX_DESC_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " ,PROC_NR_SEQUENCIA "
    str_SQL_Proc = str_SQL_Proc & " ,ATUA_TX_OPERACAO "
	str_SQL_Proc = str_SQL_Proc & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Proc = str_SQL_Proc & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Proc = str_SQL_Proc & " ) Values( "
	str_SQL_Proc = str_SQL_Proc & str_MegaProcesso & "," & int_MaxProcesso & ","
	str_SQL_Proc = str_SQL_Proc & "'" & UCase(str_NovoProcesso) & "'," & ls_Seq & ", 'I', 'XXXX', GETDATE())" 
	Set rdsNovoProcesso = Conn_db.Execute(str_SQL_Proc)

    strChave = CStr(str_MegaProcesso) & CStr(int_MaxProcesso) ' & CStr(int_SubProcesso) & CStr(int_MaxAtividade) ' & CStr(strEU)
	'call grava_log(strChave,"PROCESSO","I",0)
		
    int_Total_Gravado = int_Total_Gravado + 1
	conn_db.Close
	set conn_db = Nothing
end sub

if request("txtNovoProc1") <> "" then
   str_Seq = request("txtSeq1")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc1"),str_Seq)
end if
if request("txtNovoProc2") <> "" then
   str_Seq = request("txtSeq2")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc2"),str_Seq)
end if
if request("txtNovoProc3") <> "" then
   str_Seq = request("txtSeq3")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc3"),str_Seq)
end if
if request("txtNovoProc4") <> "" then
   str_Seq = request("txtSeq4")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc4"),str_Seq)
end if
if request("txtNovoProc5") <> "" then
   str_Seq = request("txtSeq5")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc5"),str_Seq)
end if
if request("txtNovoProc6") <> "" then
   str_Seq = request("txtSeq6")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc6"),str_Seq)
end if
if request("txtNovoProc7") <> "" then
   str_Seq = request("txtSeq7")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc7"),str_Seq)
end if
if request("txtNovoProc8") <> "" then
   str_Seq = request("txtSeq8")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc8"),str_Seq)
end if
if request("txtNovoProc9") <> "" then
   str_Seq = request("txtSeq9")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc9"),str_Seq)
end if
if request("txtNovoProc10") <> "" then
   str_Seq = request("txtSeq10")
   if str_Seq = "" then
      str_Seq = "0"
   end if
   call Grava_Novo_Processo(request("txtNovoProc10"),str_Seq)
end if

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
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
<table width="75%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%"><%'=Application("MegaProcesso")%></td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de registro gravado:</font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=int_Total_Gravado%></font></td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%"><%'=str_Opc%></td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">
        <%if str_Opc <> "1" then %>
      <table width="74%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="JavaScript:history.go(-2)"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Sele&ccedil;&atilde;o de Processo </font></td>
        </tr>
        <tr>
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
		<% end if %>
    </td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">
	 <%if str_Opc = "1" then %>
      <table width="75%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="cadas_processo.asp?txtOpc=1&selMegaProcesso=<%=str_MegaProcesso%>"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Novo Processo</font></td>
        </tr>
        <tr>
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
	  <% end if %>
    </td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
  <tr>
    <td width="14%">&nbsp;</td>
    <td width="76%">&nbsp;</td>
    <td width="10%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
