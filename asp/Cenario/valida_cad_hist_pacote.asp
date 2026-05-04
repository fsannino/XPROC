<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

pacote=request("selpacote")
versao=request("txtversao")
descricao=request("txtdescricao")
datai=request("txtdatainicio")
semana=request("txtsemana")
motivo=request("selmotivo")
obs=request("txtobs")

ssql=""
ssql="INSERT INTO PACOTE_TESTES "
ssql=ssql+"VALUES('" & pacote & "',"
ssql=ssql+"" & versao & ","
ssql=ssql+"'" & ucase(descricao) & "',"
ssql=ssql+"'" & cdate(datai) & "',"
ssql=ssql+"" & semana & ","
ssql=ssql+"'" & motivo & "',"
ssql=ssql+"'" & ucase(obs) & "',"
ssql=ssql+"'I',"
ssql=ssql+"'" & Session("CdUsuario") & "',GETDATE())"

on error resume next

'response.write ssql

db.execute(ssql)

if err.number=0 then
	aval="Registro Incluído com Sucesso!"
else
	aval="Ocorreu um erro no cadastro..." & err.description
end if
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
<script language="JavaScript" src="pupdate.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#CC3300" vlink="#CC3300" alink="#CC3300">
<form name="frm1" method="POST" action="grava_hist_pacote.asp">
  <table width="903" height="86" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" style="margin-bottom: 0">
    <tr> 
      <td height="66" colspan="2">&nbsp;</td>
      <td height="66" colspan="2">&nbsp;</td>
      <td valign="top" colspan="2" height="66"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="1">&nbsp; </td>
      <td height="20" width="1">&nbsp;</td>
      <td height="20" width="625"><table width="426" border="0" align="center">
          <tr> 
            <td width="26">&nbsp;</td>
            <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="26">&nbsp;</td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28">&nbsp;</td>
            <td width="26">&nbsp;</td>
            <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          </tr>
        </table></td>
      <td colspan="2" height="20">&nbsp;
        
      </td>
      <td height="20" width="274">&nbsp;</td>
    </tr>
  </table>
  <table width="90%" border="0">
    <tr> 
      <td width="21%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="center"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif">Cadastro de Histórico de Pacote de Testes</font></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="88%" border="0">
    <tr> 
      <td width="25%">&nbsp;</td>
      <td colspan="2"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=aval%></strong></font></td>
      <td width="8%">&nbsp;</td>
    </tr>
    <tr> 
      <td height="39">&nbsp;</td>
      <td width="10%"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
      <td width="57%"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="36">&nbsp;</td>
      <td><div align="right"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02_off.gif" width="22" height="20" border="0"></a></font></div></td>
      <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Retornar 
        para a Tela Principal</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="40">&nbsp;</td>
      <td><div align="right"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="cad_hist_pacote.asp"><img src="../../imagens/selecao_F02_off.gif" width="22" height="20" border="0"></a></font></div></td>
      <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Retornar 
        para a Tela de Cadastro de Hist&oacute;rico de Pacote de Testes</font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  </form>
</body>
</html>
