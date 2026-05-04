<%
Data_Inicio = session("data_inicio")
periodo = session("Periodo")

if session("Modo")="P" then
	modo = "Percentual"
else
	modo = "Quantitativo"
end if

%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Nova pagina 3</title>
</head>

<body>

<p align="center"><b><font size="6" face="Verdana" color="#000080">Acompanhamento de Chamados</font></b></p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="97%" id="AutoNumber1" height="298">
  <tr> 
    <td width="100%" height="290" colspan="2"> 
      <p align="center"><img border="0" src="computerkeys.jpeg" width="386" height="260">
    </td>
  </tr>
  <tr> 
    <td width="52%" height="19">&nbsp;</td>
    <td width="48%" height="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="100%" height="19" colspan="2"> 
      <p align="center"><b><font face="Verdana" size="2" color="#000080">Acompanhamento 
        de Chamados do ARS (Action Request System)</font></b>
    </td>
  </tr>
  <tr> 
    <td width="52%" height="19">&nbsp;</td>
    <td width="48%" height="19">&nbsp;</td>
  </tr>
  <tr> 
    <td width="100%" height="19" align="center" bgcolor="#CCCCCC" colspan="2"><font face="Verdana" color="#000080" size="2"><b>Configuração 
      Atual</b></font></td>
  </tr>
  <tr> 
    <td width="52%" height="19" align="center" bgcolor="#E5E5E5"><font size="2" color="#000080"><b><font face="Verdana">Data 
      Base Inicial : </font></b><font face="Verdana"><%=data_inicio%></font></font></td>
    <td width="48%" height="19" align="center" bgcolor="#E5E5E5"> 
      <p><font size="2" color="#000080"><b><font face="Verdana">Período : </font></b><font face="Verdana"><%=periodo%> 
        dias</font></font>
    </td>
  </tr>
  <tr> 
    <td width="50%" height="19" align="center" bgcolor="#E5E5E5"><b><font face="Verdana" size="2" color="#000080">Tipo</font></b><font face="Verdana" size="2" color="#000080"><b> 
      : </b><%=Session("Erro")%></font></td>
    <td width="50%" height="19" align="center" bgcolor="#E5E5E5"><b><font face="Verdana" size="2" color="#000080">Órgão 
      : </b><%=Session("Orgao")%></font></td>
  </tr>
  <tr bgcolor="#E6E6E6"> 
    <td width="100%" height="19" colspan="2"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#000099" size="2">Modo 
        de Visualiza&ccedil;&atilde;o (Perfil de Atendimento ) :</font></b> <font color="#000099" size="2"><%=modo%></font></font></div>
    </td>
  </tr>
  <tr> 
    <td width="100%" height="19" colspan="2"> 
      <p align="center" style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1" color="#800000">Para 
        alterar qualquer dos parâmetros acima, clique em CONFIGURAÇÃO, </font></b> 
      <p align="center" style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1" color="#800000">no 
        topo do menu ao lado</font></b>
    </td>
  </tr>
</table>
<p align="center" style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1" color="#800000"><%=Session("compl")%></font></b></p>
</body>

</html>