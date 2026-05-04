<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="conn_consulta.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql=""
ssql="SELECT DISTINCT TOP 100 PERCENT dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"COUNT(dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO) AS SOMA "
ssql=ssql+"FROM dbo.APOIO_LOCAL_ORGAO INNER JOIN "
ssql=ssql+"dbo.ORGAO_MENOR ON dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR "
ssql=ssql+"GROUP BY dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR, "
ssql=ssql+"dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR "
ssql=ssql+"ORDER BY dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR "

set rs=db.execute(ssql)
%>
<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;

function verifica_tecla(e)
{
if(window.event.keyCode==16)
{
alert("Tecla não permitida!");
return;
}
}

function procura()
{
window.open("procura.asp?apoio=1","_blank","width=240,height=150,history=0,scrollbars=0,titlebar=0,resizable=0,status=0")
}

</SCRIPT>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#000000" alink="#000000" onKeyDown="verifica_tecla()" link="#000000">
<form name="frm1" method="POST" action="valida_cad_orgao.asp">

  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20" colspan="2">&nbsp;</td>
      <td width="44%" height="60" colspan="2">&nbsp;</td>
      <td width="36%" valign="top" colspan="2"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center">&nbsp; 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20">&nbsp; </td>
      <td height="20">
        &nbsp;
        </td>
      <td height="20">
      <p align="right"><a href="menu.asp"><img border="0" src="../../imagens/volta_f02.gif"></a></p>
      </td>
      <td height="20"><b><font color="#000080" face="Verdana" size="2">&nbsp;Menu
        Principal</font></b> </td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
    </tr>
  </table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;
  <table border="0" width="69%">
    <tr>
      <td width="63%">
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#000080" face="Verdana" size="4">BASE
  DE APOIADORES LOCAIS</font></b></p>
        <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#000080" face="Verdana" size="1"><b>Total
        de Apoiadores Locais por Órgão Apoiado</b></font></p>
      </td>
      <td width="7%">
        <p align="right"><a href="javascript:print()"><img border="0" src="impressão.jpg" width="31" height="32"></a></td>
      <td width="37%"><font color="#000080" size="2" face="Verdana"><b>Imprimir
        Relatório</b></font></td>
    </tr>
  </table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" height="55" width="699">
    <tr>
      <td width="132" align="center" height="19"></td>
      <td width="207" align="center" height="19">
        <p align="center"><b><font color="#000080" face="Verdana" size="3"><i>Orgão
        Apoiado</i></font></b></td>
      <td width="60" align="center" height="19">
        <p align="center"><b><font color="#000080" face="Verdana" size="3"><i>Total</i></font></b></td>
      <td width="195" align="center" height="19">
        <p align="center"><b><font color="#000080" face="Verdana" size="3"><i>Órgão
        Apoiado</i></font></b></td>
      <td width="73" align="center" height="19">
        <p align="center"><b><font color="#000080" face="Verdana" size="3"><i>Total</i></font></b></td>
    </tr>
    <%
    on error resume next
    do until rs.eof=true
    %>
    <tr>
      <td width="132" align="center" height="24"></td>
      <td width="207" align="center" height="24"><font color="#000080" face="Verdana" size="1"><b><%=rs("orme_sg_org_menor")%></b></font></td>
      <td width="60" align="center" height="24" bgcolor="#C0C0C0"><font color="#000080" face="Verdana" size="1"><%=rs("soma")%></font></td>
      <%
      soma=soma+rs("soma")
      rs.movenext%>
      <td width="195" align="center" height="24"><font color="#000080" face="Verdana" size="1"><b><%=rs("orme_sg_org_menor")%></b></font></td>
      <td width="73" align="center" height="24" bgcolor="#C0C0C0"><font color="#000080" face="Verdana" size="1"><%=rs("soma")%></font></td>
    </tr>
    <%
    soma=soma+rs("soma")
    rs.movenext
    loop
    %>
  </table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
</body>
</html>
