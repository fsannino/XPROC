<<<<<<< HEAD
<%
usua=Session("CdUsuario")

%>
<html>
<head>
<script>

function foca() 
{ 
document.frm1.assunto.focus();
}

function Confirma() 
{ 
if (document.frm1.assunto.selectedIndex == 0)
{ 
alert("Especifique um ASSUNTO!");
document.frm1.assunto.focus();
return;
}
if (document.frm1.mensagem.value == "")
{ 
alert("Especifique uma MENSAGEM!");
document.frm1.mensagem.focus();
return;
}
else
{
document.frm1.submit();
}
}
</SCRIPT>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:foca()">
<form name="frm1" method="POST" action="mail.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20" colspan="2">&nbsp;</td>
      <td width="44%" height="60" colspan="2">&nbsp;</td>
      <td width="36%" valign="top" colspan="2"> 
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
      <td height="20">&nbsp; </td>
      <td height="20"><a href="javascript:Confirma()"><img border="0" src="../imagens/confirma_f02.gif" align="right"> </a> </td>
      <td height="20">&nbsp;<b><font size="2" face="Verdana" color="#330099">Enviar</font></b> </td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;
  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">X-PROC
  - Fale Conosco</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
  <table border="0" width="100%" height="292">
    <tr>
      <td width="13%" height="21"><input type="hidden" name="de" size="10" value="<%=usua%>"></td>
      <td width="14%" height="21"><b><font size="2" face="Verdana" color="#330099">Chave</font></b></td>
      <td width="73%" height="21"><b><font size="2" face="Verdana" color="#330099"><%=usua%></font></b></td>
    </tr>
    <tr>
      <td width="13%" height="9"></td>
      <td width="14%" height="9"></td>
      <td width="73%" height="9"></td>
    </tr>
    <tr>
      <td width="13%" height="19"></td>
      <td width="14%" height="19"><b><font size="2" face="Verdana" color="#330099">Assunto</font></b></td>
      <td width="73%" height="19"><select size="1" name="assunto">
          <option value="0">== Selecione ==</option>
          <option>Erros Ocorridos</option>
          <option>Melhorias</option>
          <option>Novos Relatórios</option>
          <option>Outros</option>
        </select></td>
    </tr>
    <tr>
      <td width="13%" height="24"></td>
      <td width="14%" height="24"></td>
      <td width="73%" height="24"><img border="0" src="../imagens/b021.gif" align="absmiddle">
        <font face="Verdana" size="1" color="#FF9933">Selecione o Assunto à que
        se refere seu contato</font></td>
    </tr>
    <tr>
      <td width="13%" height="227"></td>
      <td width="14%" valign="top" height="227"><b><font size="2" face="Verdana" color="#330099">Mensagem</font></b></td>
      <td width="73%" height="227"><textarea rows="11" name="mensagem" cols="60"></textarea></td>
    </tr>
    <tr>
      <td width="13%" height="1"></td>
      <td width="14%" valign="top" height="1"></td>
      <td width="73%" height="1"><font face="Verdana" size="1" color="#FF9933"><img border="0" src="../imagens/b021.gif" align="absmiddle">
        Escreva sua mensagem, descrevendo o motivo de seu contato</font></td>
    </tr>
  </table>
</form>
</body>
</html>

=======
<%
usua=Session("CdUsuario")

%>
<html>
<head>
<script>

function foca() 
{ 
document.frm1.assunto.focus();
}

function Confirma() 
{ 
if (document.frm1.assunto.selectedIndex == 0)
{ 
alert("Especifique um ASSUNTO!");
document.frm1.assunto.focus();
return;
}
if (document.frm1.mensagem.value == "")
{ 
alert("Especifique uma MENSAGEM!");
document.frm1.mensagem.focus();
return;
}
else
{
document.frm1.submit();
}
}
</SCRIPT>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="javascript:foca()">
<form name="frm1" method="POST" action="mail.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20" colspan="2">&nbsp;</td>
      <td width="44%" height="60" colspan="2">&nbsp;</td>
      <td width="36%" valign="top" colspan="2"> 
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
      <td height="20">&nbsp; </td>
      <td height="20"><a href="javascript:Confirma()"><img border="0" src="../imagens/confirma_f02.gif" align="right"> </a> </td>
      <td height="20">&nbsp;<b><font size="2" face="Verdana" color="#330099">Enviar</font></b> </td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
      <td height="20">&nbsp; </td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;
  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">X-PROC
  - Fale Conosco</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
  <table border="0" width="100%" height="292">
    <tr>
      <td width="13%" height="21"><input type="hidden" name="de" size="10" value="<%=usua%>"></td>
      <td width="14%" height="21"><b><font size="2" face="Verdana" color="#330099">Chave</font></b></td>
      <td width="73%" height="21"><b><font size="2" face="Verdana" color="#330099"><%=usua%></font></b></td>
    </tr>
    <tr>
      <td width="13%" height="9"></td>
      <td width="14%" height="9"></td>
      <td width="73%" height="9"></td>
    </tr>
    <tr>
      <td width="13%" height="19"></td>
      <td width="14%" height="19"><b><font size="2" face="Verdana" color="#330099">Assunto</font></b></td>
      <td width="73%" height="19"><select size="1" name="assunto">
          <option value="0">== Selecione ==</option>
          <option>Erros Ocorridos</option>
          <option>Melhorias</option>
          <option>Novos Relatórios</option>
          <option>Outros</option>
        </select></td>
    </tr>
    <tr>
      <td width="13%" height="24"></td>
      <td width="14%" height="24"></td>
      <td width="73%" height="24"><img border="0" src="../imagens/b021.gif" align="absmiddle">
        <font face="Verdana" size="1" color="#FF9933">Selecione o Assunto à que
        se refere seu contato</font></td>
    </tr>
    <tr>
      <td width="13%" height="227"></td>
      <td width="14%" valign="top" height="227"><b><font size="2" face="Verdana" color="#330099">Mensagem</font></b></td>
      <td width="73%" height="227"><textarea rows="11" name="mensagem" cols="60"></textarea></td>
    </tr>
    <tr>
      <td width="13%" height="1"></td>
      <td width="14%" valign="top" height="1"></td>
      <td width="73%" height="1"><font face="Verdana" size="1" color="#FF9933"><img border="0" src="../imagens/b021.gif" align="absmiddle">
        Escreva sua mensagem, descrevendo o motivo de seu contato</font></td>
    </tr>
  </table>
</form>
</body>
</html>

>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
