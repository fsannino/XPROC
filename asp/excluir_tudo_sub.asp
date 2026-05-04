<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

num_mega=request("mega")
num_processo=request("proc")
num_sub=request("sub")

sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega &" AND PROC_CD_PROCESSO="& num_processo &" AND SUPR_CD_SUB_PROCESSO=" & num_sub

ssql="DELETE FROM " & Session("PREFIXO") & "RELACAO_FINAL "
SSQL=SSQL + SQL_COMPL

db.execute(ssql)
'call grava_log(num_sub,"" & Session("PREFIXO") & "RELACAO_FINAL","D",1)

ssql="DELETE FROM " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE "
SSQL=SSQL + SQL_COMPL

db.execute(ssql)
'call grava_log(num_sub,"" & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE","D",1)


ssql="DELETE FROM " & Session("PREFIXO") & "CENARIO "
SSQL=SSQL + SQL_COMPL

db.execute(ssql)
'call grava_log(num_sub,"" & Session("PREFIXO") & "CENARIO","D",1)


ssql="DELETE FROM " & Session("PREFIXO") & "SUB_PROCESSO "
SSQL=SSQL + SQL_COMPL

db.execute(ssql)
'call grava_log(num_sub,"" & Session("PREFIXO") & "SUB_PROCESSO","D",1)
%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>


</head>


<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
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
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p><%'=ssql%></p>
<p>&nbsp;</p>
<p align="center"><b><font color="#000080" face="Verdana" size="2">Registro
Excluído com Sucesso!</font></b></p>

<p>&nbsp;</p>

<div align="center">
  <center>
  <table border="0" width="273">
    <tr>
          <td height="41" width="24"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41" width="233"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
=======
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

num_mega=request("mega")
num_processo=request("proc")
num_sub=request("sub")

sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega &" AND PROC_CD_PROCESSO="& num_processo &" AND SUPR_CD_SUB_PROCESSO=" & num_sub

ssql="DELETE FROM " & Session("PREFIXO") & "RELACAO_FINAL "
SSQL=SSQL + SQL_COMPL

db.execute(ssql)
'call grava_log(num_sub,"" & Session("PREFIXO") & "RELACAO_FINAL","D",1)

ssql="DELETE FROM " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE "
SSQL=SSQL + SQL_COMPL

db.execute(ssql)
'call grava_log(num_sub,"" & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE","D",1)


ssql="DELETE FROM " & Session("PREFIXO") & "CENARIO "
SSQL=SSQL + SQL_COMPL

db.execute(ssql)
'call grava_log(num_sub,"" & Session("PREFIXO") & "CENARIO","D",1)


ssql="DELETE FROM " & Session("PREFIXO") & "SUB_PROCESSO "
SSQL=SSQL + SQL_COMPL

db.execute(ssql)
'call grava_log(num_sub,"" & Session("PREFIXO") & "SUB_PROCESSO","D",1)
%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>


</head>


<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
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
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p><%'=ssql%></p>
<p>&nbsp;</p>
<p align="center"><b><font color="#000080" face="Verdana" size="2">Registro
Excluído com Sucesso!</font></b></p>

<p>&nbsp;</p>

<div align="center">
  <center>
  <table border="0" width="273">
    <tr>
          <td height="41" width="24"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41" width="233"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
