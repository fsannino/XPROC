<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_bug=0
str_bug=request("selBug")

set rs=db.execute("SELECT DISTINCT BUG_CD_BUG, BUG_MEGA_PROCESSO FROM " & Session("PREFIXO") & "BUG_IMPORT ORDER BY BUG_CD_BUG")

select case request("order")
case 1
	ORDENA="MODU_TX_MODULO"
case 2
	ORDENA="ATCA_TX_ATIVIDADE"
case 3
	ORDENA="TRAN_TX_TRANSACAO"
case 4
	ORDENA="BUG_TX_PROBLEMA"
case else
	ORDENA="BUG_CD_BUG"
END SELECT

if str_bug<>0 then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "BUG_IMPORT WHERE BUG_CD_BUG=" & str_bug &" ORDER BY " & ORDENA)
else
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "BUG_IMPORT WHERE BUG_CD_BUG=0")
end if

%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" action="rel_bug.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>
<p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3"> Erros de Importação
de Dados</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="96%" height="54">
  <tr>
    <td width="35%" height="1"></td>
    <td width="65%" height="1"><b><font face="Verdana" size="2" color="#330099">Selecione
      o Identificador da Importação</font></b></td>
  </tr>
  <tr>
    <td width="35%" height="25"></td>
    <td width="65%" height="25"><select size="1" name="selBug" onchange="javascript:submit()">
        <option value="0">== Selecione o Identificador da Importação ==</option>
       <%do until rs.eof=true
       if trim(str_bug)=trim(rs("BUG_CD_BUG")) then
       %>
       <option selected value=<%=RS("BUG_CD_BUG")%>><%=RS("BUG_CD_BUG")%> - <%=RS("BUG_MEGA_PROCESSO")%></option>
		<%ELSE%>
       <option value=<%=RS("BUG_CD_BUG")%>><%=RS("BUG_CD_BUG")%> - <%=RS("BUG_MEGA_PROCESSO")%></option>
		<%
		end if
		rs.movenext
		loop
		%>	        
       </select></td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
<div align="center">
  <center>
<table border="0" width="815" cellspacing="3">
<%if rs1.eof=false then%>
  <tr>
    <td width="22" bgcolor="#CAE2E3"><font face="Verdana" size="1"><b>ID</b></font></td>
    <td width="181" bgcolor="#CAE2E3"><font face="Verdana" size="1"><b><a href="rel_bug.asp?selBug=<%=str_bug%>&order=1">Agrupamento
      das Atividade</a></b></font></td>
    <td width="209" bgcolor="#CAE2E3"><font face="Verdana" size="1"><a href="rel_bug.asp?selBug=<%=str_bug%>&order=2"><b>Atividade</b></a></font></td>
    <td width="72" bgcolor="#CAE2E3"><font face="Verdana" size="1"><a href="rel_bug.asp?selBug=<%=str_bug%>&order=3"><b>Transação</b></a></font></td>
    <td width="287" bgcolor="#CAE2E3"><font face="Verdana" size="1"><a href="rel_bug.asp?selBug=<%=str_bug%>&order=4"><b>Problema
      Ocorrido</b></a></font></td>
  </tr>
  <%end if%>
  <%do until rs1.eof=true%>
  <tr>
    <td width="22"><font face="Verdana" size="1"><%=rs1("BUG_CD_BUG")%></font></td>
    <td width="181"><font face="Verdana" size="1"><%=rs1("MODU_TX_MODULO")%></font></td>
    <td width="209"><font face="Verdana" size="1"><%=rs1("ATCA_TX_ATIVIDADE")%></font></td>
    <td width="72"><font face="Verdana" size="1"><%=TRIM(rs1("TRAN_TX_TRANSACAO"))%></font></td>
    <td width="287"><font face="Verdana" size="1"><%=rs1("BUG_TX_PROBLEMA")%></font></td>
  </tr>
  <%
  rs1.movenext
  loop
  %>
</table>

  </center>
</div>

</body>

</html>
=======
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_bug=0
str_bug=request("selBug")

set rs=db.execute("SELECT DISTINCT BUG_CD_BUG, BUG_MEGA_PROCESSO FROM " & Session("PREFIXO") & "BUG_IMPORT ORDER BY BUG_CD_BUG")

select case request("order")
case 1
	ORDENA="MODU_TX_MODULO"
case 2
	ORDENA="ATCA_TX_ATIVIDADE"
case 3
	ORDENA="TRAN_TX_TRANSACAO"
case 4
	ORDENA="BUG_TX_PROBLEMA"
case else
	ORDENA="BUG_CD_BUG"
END SELECT

if str_bug<>0 then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "BUG_IMPORT WHERE BUG_CD_BUG=" & str_bug &" ORDER BY " & ORDENA)
else
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "BUG_IMPORT WHERE BUG_CD_BUG=0")
end if

%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" action="rel_bug.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>
<p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3"> Erros de Importação
de Dados</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="96%" height="54">
  <tr>
    <td width="35%" height="1"></td>
    <td width="65%" height="1"><b><font face="Verdana" size="2" color="#330099">Selecione
      o Identificador da Importação</font></b></td>
  </tr>
  <tr>
    <td width="35%" height="25"></td>
    <td width="65%" height="25"><select size="1" name="selBug" onchange="javascript:submit()">
        <option value="0">== Selecione o Identificador da Importação ==</option>
       <%do until rs.eof=true
       if trim(str_bug)=trim(rs("BUG_CD_BUG")) then
       %>
       <option selected value=<%=RS("BUG_CD_BUG")%>><%=RS("BUG_CD_BUG")%> - <%=RS("BUG_MEGA_PROCESSO")%></option>
		<%ELSE%>
       <option value=<%=RS("BUG_CD_BUG")%>><%=RS("BUG_CD_BUG")%> - <%=RS("BUG_MEGA_PROCESSO")%></option>
		<%
		end if
		rs.movenext
		loop
		%>	        
       </select></td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
<div align="center">
  <center>
<table border="0" width="815" cellspacing="3">
<%if rs1.eof=false then%>
  <tr>
    <td width="22" bgcolor="#CAE2E3"><font face="Verdana" size="1"><b>ID</b></font></td>
    <td width="181" bgcolor="#CAE2E3"><font face="Verdana" size="1"><b><a href="rel_bug.asp?selBug=<%=str_bug%>&order=1">Agrupamento
      das Atividade</a></b></font></td>
    <td width="209" bgcolor="#CAE2E3"><font face="Verdana" size="1"><a href="rel_bug.asp?selBug=<%=str_bug%>&order=2"><b>Atividade</b></a></font></td>
    <td width="72" bgcolor="#CAE2E3"><font face="Verdana" size="1"><a href="rel_bug.asp?selBug=<%=str_bug%>&order=3"><b>Transação</b></a></font></td>
    <td width="287" bgcolor="#CAE2E3"><font face="Verdana" size="1"><a href="rel_bug.asp?selBug=<%=str_bug%>&order=4"><b>Problema
      Ocorrido</b></a></font></td>
  </tr>
  <%end if%>
  <%do until rs1.eof=true%>
  <tr>
    <td width="22"><font face="Verdana" size="1"><%=rs1("BUG_CD_BUG")%></font></td>
    <td width="181"><font face="Verdana" size="1"><%=rs1("MODU_TX_MODULO")%></font></td>
    <td width="209"><font face="Verdana" size="1"><%=rs1("ATCA_TX_ATIVIDADE")%></font></td>
    <td width="72"><font face="Verdana" size="1"><%=TRIM(rs1("TRAN_TX_TRANSACAO"))%></font></td>
    <td width="287"><font face="Verdana" size="1"><%=rs1("BUG_TX_PROBLEMA")%></font></td>
  </tr>
  <%
  rs1.movenext
  loop
  %>
</table>

  </center>
</div>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
