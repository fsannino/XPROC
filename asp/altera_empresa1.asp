<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

str_emp=request("selEmp")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE WHERE EMPR_CD_NR_EMPRESA="& str_emp)
%>
<html>
<head>
<script>
function Confirma() 
{ 
if (document.frm1.DescEmpresa.value == "")
     { 
	 alert("A Descrição de uma Empresa é obrigatório!");
     document.frm1.DescEmpresa.focus();
     return;
     }
 }
</SCRIPT>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK REL="SHORTCUT ICON" href="http://regina/imagens/Wrench.ico">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="altera_empresa2.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="152" height="20" colspan="2">&nbsp;</td>
      <td width="337" height="60" colspan="3">&nbsp;</td>
      <td width="278" valign="top" colspan="2"> 
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
      <td height="20" width="0%">&nbsp; </td>
      <td height="20" width="9%">&nbsp;</td>
      <td height="20" width="3%"> 
        <p align="center"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:submit()">
      </td>
      <td height="20" width="36%"><font face="Verdana" size="2" color="#330099"><b>Enviar</b></font></td>
      <td colspan="2" height="20" width="69">&nbsp;</td>
      <td height="20" width="44%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="50%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Alteração 
        de Empresa / Unidade</font></td>
      <td width="26%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2">&nbsp;</font></b></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2">Empresa</font></b></td>
      <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <p> 
          <input type="text" name="DescEmpresa" size="75" value="<%=rs("EMPR_TX_NOME_EMPRESA")%>">
        </p>
        </font></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%"></td>
      <td width="13%"></td>
      <td width="63%">&nbsp; 
        <input type="hidden" name="CodEmpresa" size="7" value="<%=rs("EMPR_CD_NR_EMPRESA")%>">
      </td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8%">&nbsp;</td>
      <td width="13%">&nbsp;</td>
      <td width="63%">&nbsp;</td>
    </tr>
  </table> 
</form>
</body>
</html>
