<%
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body topmargin="0" leftmargin="0" link="#000080" vlink="#000080" alink="#000080">
<form>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="89%" id="AutoNumber2" height="514">
           <tr>
                      <td width="100%" height="68" valign="top" colspan="2">
<img border="0" src="topo.jpg"></td>
           </tr>
           <tr>
                      <td width="13%" height="445" valign="top"><img border="0" src="lado.jpg" width="83" height="429"></td>
                      <td width="87%" height="445" valign="top">
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="93%" id="AutoNumber3" height="96">
                         <tr>
                                    <td width="40%" height="130" align="center" colspan="2"><img border="0" src="mult_c.jpg" align="right"></td>
                                    <td width="60%" height="130" align="left"><font face="Verdana" color="#800000"><b>CONSULTAS</b></font></td>
                         </tr>
                         <tr>
                                    <td width="35%" height="41" align="center">&nbsp;</td>
                                    <td width="5%" height="41" align="center"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
                                    <td width="60%" height="41" align="left"><b><font face="Verdana" size="2"><a href="selecao_consulta.asp?tipo=1">MULTIPLICADOR X CURSO</a></font></b></td>
                         </tr>
                         <tr>
                                    <td width="35%" height="43" align="center">&nbsp;</td>
                                    <td width="5%" height="43" align="center"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
                                    <td width="60%" height="43" align="left"><b><font face="Verdana" size="2"><a href="selecao_consulta.asp?tipo=2">CURSO X MULTIPLICADOR</a></font></b></td>
                         </tr>
                         <%
                         'if Session("Acesso")=9 then
                         %>
                         <tr>
                                    <td width="35%" height="43" align="center">&nbsp;</td>
                                    <td width="5%" height="43" align="center"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
                                    
            <td width="60%" height="43" align="left"><b><font face="Verdana" size="2"><a href="cons_geral.asp">CONSULTA 
              DE TODOS OS MAPEADOS POR ÓRGÃO</a></font></b></td>
                         </tr>
                         <tr>
                                    <td width="35%" height="43" align="center">&nbsp;</td>
                                    <td width="5%" height="43" align="center"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
                                    <td width="60%" height="43" align="left"><b><font face="Verdana" size="2"><a href="cons_demons.asp?selOrgao=88">CONSULTA DEMONSTRATIVA POR ÓRGÃO</a></font></b></td>
                         </tr>
                         <%
                         'end if
                         %>
                         <tr>
                                    <td width="100%" height="45" align="center" colspan="3">&nbsp;<p><a href="javascript:history.go(-1)"><img border="0" src="voltar.gif"></a></td>
                         </tr>
                         </table>
                      </td>
           </tr>
</table>
</form>
</body>

</html>

<script>
document.title = 'Indicação de Multiplicadores'
</script>