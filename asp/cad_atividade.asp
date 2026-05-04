 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
set rs=db.execute("SELECT MAX(ATCA_NR_SEQUENCIA) AS CODIGO FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA")

IF ISNULL(RS("CODIGO")) THEN
	VALOR=10
ELSE
	VALOR=RS("CODIGO")+10
END IF
	

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function Confirma() 
{ 
if (document.form1.CadAtividade.value == "")
     { 
	 alert("O campo DESCRIÇÃO DA ATIVIDADE deve ser preenchido!");
     document.form1.CadAtividade.focus();
     return;
     }
	 else
     {
	  document.form1.submit();
	 }
}
function Limpa(){
	document.form1.reset();
}

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="form1" method="POST" action="valida_cad_atividade.asp">
  <table width="773" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td height="66" colspan="2">&nbsp;</td>
      <td height="66" colspan="2">&nbsp;</td>
      <td valign="top" colspan="2" height="66"> 
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
      <td height="20" width="111">&nbsp; </td>
      <td height="20" width="30"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
      <td height="20" width="213"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Enviar</b></font></td>
      <td colspan="2" height="20">
        <div align="right"><a href="javascript:Limpa()"><img src="../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></div>
      </td>
      <td height="20" width="334"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Limpa</b></font></td>
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
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Cadastro 
        de Atividade</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="5%">&nbsp;</td>
      <td width="22%">&nbsp;</td>
      <td width="36%">&nbsp;</td>
      <td width="37%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="22%">&nbsp;</td>
      <td width="36%">&nbsp;</td>
      <td width="37%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="22%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descrição 
        da Atividade</b></font></td>
      <td width="36%"> 
        <input type="text" name="CadAtividade" size="45" VALUE="">
      </td>
      <td width="37%">&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Sequência 
        : </b></font> 
        <input type="text" name="CadSequencia" size="7" value="<%=VALOR%>">
      </td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="22%">&nbsp;</td>
      <td width="36%">&nbsp;</td>
      <td width="37%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="5%"></td>
      <td width="22%"></td>
      <td width="1%"></td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="22%">&nbsp;</td>
      <td width="36%">&nbsp;</td>
      <td width="37%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="22%">&nbsp;</td>
      <td width="36%">&nbsp;</td>
      <td width="37%">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
