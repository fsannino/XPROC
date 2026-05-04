<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE WHERE MEPR_CD_MEGA_PROCESSO="& request.querystring("selMegaProcesso")& "ORDER BY ATIV_TX_DESC_ATIVIDADE")
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
	window.location.href='relacao3.asp?selMegaProcesso='+document.formulario.selMegaProcesso.value+'&selAtividade='+document.formulario.selAtividade.value
}
//-->
</script>

<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFF00">
  <tr bgcolor="#330099">
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="85%" border="0" align="right" cellpadding="0" cellspacing="1" bgcolor="#0000CC">
        <tr>
          <td bgcolor="#330099">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="JavaScript:history.back()"><font color="#FFFFFF">Volta</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="JavaScript:history.forward()"><font color="#FFFFFF">Proximo</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><font color="#FFFFFF">Favorito</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Para</font></b></font>
            </div>
          </td>
        </tr>
        <tr>
          <td bgcolor="#330099" height="12">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Impr</font></b></font>
            </div>
          </td>
          <td bgcolor="#330099" height="12">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="JavaScript:history.go()"><font color="#FFFFFF">Atualiza</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099" height="12">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="file:///E:/xproc/index.asp"><font color="#FFFFFF">Inicial</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099" height="12">
            <div align="center">
            </div>
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
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0"><img border="0" src="../Imagens/topo_relaciona.gif" width="450" height="51"></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<form method="POST" action="resultado_consulta.asp" name="formulario">
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">Selecione o Mega-Processo&nbsp;&nbsp;
  <select size="1" name="selMegaProcesso" onChange="MM_goToURL('parent','relacao2.asp?txtOpc=1&amp;selMegaProcesso=');return document.MM_returnValue">
  <option value=0>== Selecione ==</option>
    <%
  	 DO UNTIL RS.EOF=TRUE
  	 
  	 response.write request.querystring("selMegaProcesso")
  	 response.write rs("MEPR_CD_MEGA_PROCESSO")
  	 
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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">Selecione
  a Atividade&nbsp;&nbsp;&nbsp; <select size="1" name="selAtividade" onChange="javascript:goToURL2()">
    <option value="0">== Selecione ==</option>
    <%
  	 DO UNTIL RS2.EOF=TRUE
    %>  
    <option value=<%=rs2("ATIV_CD_ATIVIDADE")%>><%=rs2("ATIV_TX_DESC_ATIVIDADE")%></option>
    <%
    RS2.MOVENEXT
    LOOP
    %>
  </select></font></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" width="35%">
    <tr>
      <td width="51%">  </td>
      <td width="49%"></td>
    </tr>
  </table>
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

set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE WHERE MEPR_CD_MEGA_PROCESSO="& request.querystring("selMegaProcesso")& "ORDER BY ATIV_TX_DESC_ATIVIDADE")
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
	window.location.href='relacao3.asp?selMegaProcesso='+document.formulario.selMegaProcesso.value+'&selAtividade='+document.formulario.selAtividade.value
}
//-->
</script>

<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFF00">
  <tr bgcolor="#330099">
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="85%" border="0" align="right" cellpadding="0" cellspacing="1" bgcolor="#0000CC">
        <tr>
          <td bgcolor="#330099">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="JavaScript:history.back()"><font color="#FFFFFF">Volta</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="JavaScript:history.forward()"><font color="#FFFFFF">Proximo</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><font color="#FFFFFF">Favorito</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Para</font></b></font>
            </div>
          </td>
        </tr>
        <tr>
          <td bgcolor="#330099" height="12">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Impr</font></b></font>
            </div>
          </td>
          <td bgcolor="#330099" height="12">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="JavaScript:history.go()"><font color="#FFFFFF">Atualiza</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099" height="12">
            <div align="center">
              <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="file:///E:/xproc/index.asp"><font color="#FFFFFF">Inicial</font></a></b></font>
            </div>
          </td>
          <td bgcolor="#330099" height="12">
            <div align="center">
            </div>
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
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0"><img border="0" src="../Imagens/topo_relaciona.gif" width="450" height="51"></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<form method="POST" action="resultado_consulta.asp" name="formulario">
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">Selecione o Mega-Processo&nbsp;&nbsp;
  <select size="1" name="selMegaProcesso" onChange="MM_goToURL('parent','relacao2.asp?txtOpc=1&amp;selMegaProcesso=');return document.MM_returnValue">
  <option value=0>== Selecione ==</option>
    <%
  	 DO UNTIL RS.EOF=TRUE
  	 
  	 response.write request.querystring("selMegaProcesso")
  	 response.write rs("MEPR_CD_MEGA_PROCESSO")
  	 
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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">Selecione
  a Atividade&nbsp;&nbsp;&nbsp; <select size="1" name="selAtividade" onChange="javascript:goToURL2()">
    <option value="0">== Selecione ==</option>
    <%
  	 DO UNTIL RS2.EOF=TRUE
    %>  
    <option value=<%=rs2("ATIV_CD_ATIVIDADE")%>><%=rs2("ATIV_TX_DESC_ATIVIDADE")%></option>
    <%
    RS2.MOVENEXT
    LOOP
    %>
  </select></font></p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" width="35%">
    <tr>
      <td width="51%">  </td>
      <td width="49%"></td>
    </tr>
  </table>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
