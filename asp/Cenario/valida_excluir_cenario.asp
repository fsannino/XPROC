<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

codigo=request("ID")

if len(request("ID2"))>0 THEN
	codigo=request("ID2")
END IF

tem=0

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & codigo & "'"
set rst_Cenario=db.execute(ssql)
if rst_Cenario.EOF then
   response.redirect "msg.asp?pOpt=0&txtCenario=" & codigo   	
end if

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO_SEGUINTE='" & codigo & "'"

set rs=db.execute(ssql)
if rs.eof=true then
   on error resume next
   ssql=""
   ssql="DELETE FROM " & Session("PREFIXO") & "HISTORICO_CENARIO WHERE CENA_CD_CENARIO='" & codigo & "'"
   db.execute(ssql)
   ssql="" 
   ssql="DELETE FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & codigo & "'"
   db.execute(ssql)
   ssql="DELETE FROM " & Session("PREFIXO") & "CURSO_CENARIO WHERE CENA_CD_CENARIO='" & codigo & "'"
   db.execute(ssql)
   SSQL=""
   ssql="DELETE FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & codigo & "'"
   db.execute(ssql)
else
   tem=1
end if

%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body topmargin="0" leftmargin="0">
<form method="POST" action="" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
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
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp;</font></p>

  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Exclusão
  de Cenário</font></p>

  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<table border="0" width="84%" height="123">
  <%if err.number=0 and tem=0 then
  valor_cod="O registro foi excluído com Sucesso"
  %>
  <tr>
    <td width="32%" height="21">&nbsp;</td>
    <td width="70%" colspan="3" height="21"><font face="Verdana" size="2" color="#330099"><b><%=valor_cod%></b></font></td>
  </tr>
  <%
  else
  if tem=1 then
	  valor_cod=" Este Cenário é continuação de um outro cenário. Não é possível excluí-lo"
  else
	  valor_cod=" Não foi possível excluir o registro - " & err.description
  end if
  %>
  <tr>
    <td width="32%" height="21">&nbsp;</td>
      <td width="70%" colspan="3" height="21"><b><font face="Verdana" size="2" color="#800000"><%=valor_cod%></font></b></td>
  </tr>
  <tr>
    <td width="32%" height="21">&nbsp;</td>
    <td width="70%" colspan="3" height="21">&nbsp;</td>
  </tr>
  <%end if%>
  <tr>
    <td width="32%" height="22">&nbsp;</td>
    <td width="7%" height="22">&nbsp;</td>
    <td width="7%" height="22">
      <p align="right"><font face="Verdana" size="2" color="#330099"><a href="../../indexA.asp"><img src="selecao_F02.gif" width="22" height="20" border="0"></a></font></td>
    <td width="56%" height="22"><font face="Verdana" size="2" color="#330099">Volta para a tela principal</font></td>
  </tr>
  <tr>
    <td width="32%" height="1"></td>
    <td width="7%" height="1"></td>
    <td width="7%" height="1">
      <p align="right"><font face="Verdana" size="2" color="#330099">
      <%IF TEM=1 THEN%>
      <a href="gera_rel_relac_cenario.asp?id=<%=codigo%>&amp;selVis=2"><img src="selecao_F02.gif" width="22" height="20" border="0"></a></font></td>
    <td width="56%" height="1"><font face="Verdana" size="2" color="#330099">Exibir
      Relação com Cenários</font>
      <%END IF%>
      </td>
  </tr>
  <tr>
    <td width="32%" height="1">&nbsp;</td>
    <td width="7%" height="1">&nbsp;</td>
    <td width="7%" height="1">
      <p align="right"><font face="Verdana" size="2" color="#330099"><a href="excluir_cenario.asp"><img src="selecao_F02.gif" width="22" height="20" border="0"></a></font></td>
    <td width="56%" height="1"><font face="Verdana" size="2" color="#330099">Volta para a
      tela de exclusão de Cenário</font></td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</body>
</html>
