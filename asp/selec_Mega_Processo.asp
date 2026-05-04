<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

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
<title>SINERGIA # XPROC # Processos de Negócio</title>
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
<form name="frm1" method="post" action="selec_Processo.asp">
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
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"><a href="javascript:Limpa()"><img src="../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">limpa</font></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="90%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Mega 
        Processo</font></td>
      <td width="73%"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','selec_Processo.asp?txtOpc=1&amp;selMegaProcesso=',this);return document.MM_returnValue">
          <option value="0" selected>Selecione um Mega Processo</option>
          <%Set rdsMegaProcesso = Conn_db.Execute(str_SQL_MegaProc)
While (NOT rdsMegaProcesso.EOF)%>
          <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>" ><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
          <%
  rdsMegaProcesso.MoveNext()
Wend
If (rdsMegaProcesso.CursorType > 0) Then
  rdsMegaProcesso.MoveFirst
Else
  rdsMegaProcesso.Requery
End If
rdsMegaProcesso.close
set rdsMegaProcesso = Nothing
%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp; </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

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
<title>SINERGIA # XPROC # Processos de Negócio</title>
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
<form name="frm1" method="post" action="selec_Processo.asp">
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
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"><a href="javascript:Limpa()"><img src="../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">limpa</font></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="90%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Mega 
        Processo</font></td>
      <td width="73%"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','selec_Processo.asp?txtOpc=1&amp;selMegaProcesso=',this);return document.MM_returnValue">
          <option value="0" selected>Selecione um Mega Processo</option>
          <%Set rdsMegaProcesso = Conn_db.Execute(str_SQL_MegaProc)
While (NOT rdsMegaProcesso.EOF)%>
          <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>" ><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
          <%
  rdsMegaProcesso.MoveNext()
Wend
If (rdsMegaProcesso.CursorType > 0) Then
  rdsMegaProcesso.MoveFirst
Else
  rdsMegaProcesso.Requery
End If
rdsMegaProcesso.close
set rdsMegaProcesso = Nothing
%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp; </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="73%">&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
