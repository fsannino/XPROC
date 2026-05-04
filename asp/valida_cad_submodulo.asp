<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")
desc=request("txtDesc")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_MODULO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)

if rs.eof=true then
	codigo=1
else
	set temp=db.execute("SELECT MAX(SUMO_NR_SEQUENCIA)AS COD FROM " & Session("PREFIXO") & "SUB_MODULO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)
	codigo=temp("COD")+1
END IF

ssql=""
ssql="INSERT INTO " & Session("PREFIXO") & "SUB_MODULO "
ssql=ssql & "VALUES('" & UCASE(DESC) & "', "
ssql=ssql & MEGA & ", "
ssql=ssql & "'I', "
ssql=ssql & codigo & ", "
ssql=ssql & "'" & Session("CdUsuario") & "', "
ssql=ssql & "GETDATE())"

'response.write ssql

db.execute(ssql)
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
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="">
        <input type="hidden" name="txtOpc" value="<%=str_OPC%>">
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
          <td width="27"></td>
          <td width="50"></td>
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
      <td width="20%" colspan="2"><%'=Session("MegaProcesso")%></td>
      <td width="70%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2"><%'=str_Opc%></td>
      <td width="70%"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Cadastro
        de Sub-Módulo</font></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2"><%'=str_SQL_Max_Seq_Proc%></td>
      <td width="70%"> <%'=str_MegaProcesso%> 
        <%'=Session("Conn_String_Cogest_Gravacao")%>
        <%'=ls_int_MaxProcesso%>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2">
      </td>
      <td width="70%"> 
      <font face="Verdana" color="#330099" size="2"><b>O Registro foi cadastrado
      com Sucesso!</b></font> 
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2"></td>
      <td width="70%"></td>
    </tr>
    <tr> 
      <td width="10%" align="right"></td>
      <td width="10%" align="right"></td>
      <td width="10%" align="right"><a href="cad_submodulo.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
      <td width="70%">&nbsp; <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Retornar
        para a tela de cadastro de Sub-Módulo</font></td>
    </tr>
    <tr> 
      <td width="10%" align="right"></td>
      <td width="10%" align="right"></td>
      <td width="10%" align="right"></td>
      <td width="70%"></td>
    </tr>
    <tr> 
      <td width="10%" align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>&nbsp;</b></font></td>
      <td width="10%" align="right"></td>
      <td width="10%" align="right"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
      <td width="70%">&nbsp; <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Retornar
        para a Tela Principal</font></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2">&nbsp;</td>
      <td width="70%">&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")
desc=request("txtDesc")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_MODULO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)

if rs.eof=true then
	codigo=1
else
	set temp=db.execute("SELECT MAX(SUMO_NR_SEQUENCIA)AS COD FROM " & Session("PREFIXO") & "SUB_MODULO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)
	codigo=temp("COD")+1
END IF

ssql=""
ssql="INSERT INTO " & Session("PREFIXO") & "SUB_MODULO "
ssql=ssql & "VALUES('" & UCASE(DESC) & "', "
ssql=ssql & MEGA & ", "
ssql=ssql & "'I', "
ssql=ssql & codigo & ", "
ssql=ssql & "'" & Session("CdUsuario") & "', "
ssql=ssql & "GETDATE())"

'response.write ssql

db.execute(ssql)
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
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="">
        <input type="hidden" name="txtOpc" value="<%=str_OPC%>">
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
          <td width="27"></td>
          <td width="50"></td>
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
      <td width="20%" colspan="2"><%'=Session("MegaProcesso")%></td>
      <td width="70%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2"><%'=str_Opc%></td>
      <td width="70%"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Cadastro
        de Sub-Módulo</font></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2"><%'=str_SQL_Max_Seq_Proc%></td>
      <td width="70%"> <%'=str_MegaProcesso%> 
        <%'=Session("Conn_String_Cogest_Gravacao")%>
        <%'=ls_int_MaxProcesso%>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2">
      </td>
      <td width="70%"> 
      <font face="Verdana" color="#330099" size="2"><b>O Registro foi cadastrado
      com Sucesso!</b></font> 
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2"></td>
      <td width="70%"></td>
    </tr>
    <tr> 
      <td width="10%" align="right"></td>
      <td width="10%" align="right"></td>
      <td width="10%" align="right"><a href="cad_submodulo.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
      <td width="70%">&nbsp; <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Retornar
        para a tela de cadastro de Sub-Módulo</font></td>
    </tr>
    <tr> 
      <td width="10%" align="right"></td>
      <td width="10%" align="right"></td>
      <td width="10%" align="right"></td>
      <td width="70%"></td>
    </tr>
    <tr> 
      <td width="10%" align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>&nbsp;</b></font></td>
      <td width="10%" align="right"></td>
      <td width="10%" align="right"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
      <td width="70%">&nbsp; <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Retornar
        para a Tela Principal</font></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%" colspan="2">&nbsp;</td>
      <td width="70%">&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
