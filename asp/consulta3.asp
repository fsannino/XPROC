<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& request.querystring("selMegaProcesso")& "ORDER BY PROC_TX_DESC_PROCESSO")

set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& request.querystring("selMegaProcesso")& " AND PROC_CD_PROCESSO="& request.querystring("selProcesso") &" ORDER BY SUPR_TX_DESC_SUB_PROCESSO")
%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.formulario.selMegaProcesso.value+"'");
}
function goToURL2() {
	window.location.href='consulta3.asp?selMegaProcesso='+document.formulario.selMegaProcesso.value+'&selProcesso='+document.formulario.selProcesso.value
}
//-->
</script>

<body topmargin="0" leftmargin="0">
<form method="POST" action="resultado_consulta.asp" name="formulario">

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
          <td width="26"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:submit()"></td>
          <td width="50"><b><font size="2" color="#330099" face="Verdana">Consultar</font></b></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana" color="#330099">Consulta
Relação Mega-Processo x Processo x Sub-Processo x Atividade x Transação</font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="2"><b>Selecione o Mega-Processo</b></font><font face="Arial" size="2">&nbsp;&nbsp;
  <select size="1" name="selMegaProcesso" onChange="MM_goToURL('parent','consulta2.asp?txtOpc=1&amp;selMegaProcesso=');return document.MM_returnValue">
  <option value=0>== Selecione ==</option>
    <%
  	 DO UNTIL RS.EOF=TRUE
  	 IF trim(request.querystring("selMegaProcesso"))= trim(rs("MEPR_CD_MEGA_PROCESSO"))THEN
    %>  
    <option selected value=<%=rs("MEPR_CD_MEGA_PROCESSO")%>><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
    <%
    ELSE
    %>
    <option value=<%=rs("MEPR_CD_MEGA_PROCESSO")%>><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
    <%
    END IF
    RS.MOVENEXT
    LOOP
    %>
  </select></font></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="2"><b>Selecione o Processo</b></font><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;
  <select size="1" name="selProcesso" onChange="javascript:goToURL2()">
    <option value=0>== Selecione ==</option>
    <%
  	 DO UNTIL RS2.EOF=TRUE
   	 IF trim(request.querystring("selProcesso"))= trim(rs2("PROC_CD_PROCESSO"))THEN
    %>  
    <option selected value=<%=rs2("PROC_CD_PROCESSO")%>><%=rs2("PROC_TX_DESC_PROCESSO")%></option>
    <%
    else
    %>
    <option value=<%=rs2("PROC_CD_PROCESSO")%>><%=rs2("PROC_TX_DESC_PROCESSO")%></option>
    <%
    END IF
    RS2.MOVENEXT
    LOOP
    %>
  </select></font></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="2"><b>Selecione
  o Sub-Processo</b></font><font face="Arial" size="2"> <select size="1" name="selSubProcesso">
    <option value=0>== Selecione ==</option>
    <%
    DO UNTIL RS3.EOF=TRUE
    %>
    <option value=<%=rs3("SUPR_CD_SUB_PROCESSO")%>><%=rs3("SUPR_TX_DESC_SUB_PROCESSO")%></option>
    <%
    RS3.MOVENEXT
    LOOP
    %>
  </select></font></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>

</html>
=======
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& request.querystring("selMegaProcesso")& "ORDER BY PROC_TX_DESC_PROCESSO")

set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& request.querystring("selMegaProcesso")& " AND PROC_CD_PROCESSO="& request.querystring("selProcesso") &" ORDER BY SUPR_TX_DESC_SUB_PROCESSO")
%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.formulario.selMegaProcesso.value+"'");
}
function goToURL2() {
	window.location.href='consulta3.asp?selMegaProcesso='+document.formulario.selMegaProcesso.value+'&selProcesso='+document.formulario.selProcesso.value
}
//-->
</script>

<body topmargin="0" leftmargin="0">
<form method="POST" action="resultado_consulta.asp" name="formulario">

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
          <td width="26"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:submit()"></td>
          <td width="50"><b><font size="2" color="#330099" face="Verdana">Consultar</font></b></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana" color="#330099">Consulta
Relação Mega-Processo x Processo x Sub-Processo x Atividade x Transação</font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="2"><b>Selecione o Mega-Processo</b></font><font face="Arial" size="2">&nbsp;&nbsp;
  <select size="1" name="selMegaProcesso" onChange="MM_goToURL('parent','consulta2.asp?txtOpc=1&amp;selMegaProcesso=');return document.MM_returnValue">
  <option value=0>== Selecione ==</option>
    <%
  	 DO UNTIL RS.EOF=TRUE
  	 IF trim(request.querystring("selMegaProcesso"))= trim(rs("MEPR_CD_MEGA_PROCESSO"))THEN
    %>  
    <option selected value=<%=rs("MEPR_CD_MEGA_PROCESSO")%>><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
    <%
    ELSE
    %>
    <option value=<%=rs("MEPR_CD_MEGA_PROCESSO")%>><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
    <%
    END IF
    RS.MOVENEXT
    LOOP
    %>
  </select></font></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="2"><b>Selecione o Processo</b></font><font face="Arial" size="2">&nbsp;&nbsp;&nbsp;
  <select size="1" name="selProcesso" onChange="javascript:goToURL2()">
    <option value=0>== Selecione ==</option>
    <%
  	 DO UNTIL RS2.EOF=TRUE
   	 IF trim(request.querystring("selProcesso"))= trim(rs2("PROC_CD_PROCESSO"))THEN
    %>  
    <option selected value=<%=rs2("PROC_CD_PROCESSO")%>><%=rs2("PROC_TX_DESC_PROCESSO")%></option>
    <%
    else
    %>
    <option value=<%=rs2("PROC_CD_PROCESSO")%>><%=rs2("PROC_TX_DESC_PROCESSO")%></option>
    <%
    END IF
    RS2.MOVENEXT
    LOOP
    %>
  </select></font></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="2"><b>Selecione
  o Sub-Processo</b></font><font face="Arial" size="2"> <select size="1" name="selSubProcesso">
    <option value=0>== Selecione ==</option>
    <%
    DO UNTIL RS3.EOF=TRUE
    %>
    <option value=<%=rs3("SUPR_CD_SUB_PROCESSO")%>><%=rs3("SUPR_TX_DESC_SUB_PROCESSO")%></option>
    <%
    RS3.MOVENEXT
    LOOP
    %>
  </select></font></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
