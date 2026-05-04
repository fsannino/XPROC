<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")

ssql="SELECT * FROM " & Session("PREFIXO") & "TRANSACAO_MEGA WHERE MEPR_CD_MEGA_PROCESSO=" & mega & " ORDER BY TRAN_CD_TRANSACAO"

set rs=db.execute(ssql)
%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="gera_rel_megaemp.asp">
              <input type="hidden" name="txtEmpSelecionada"><input type="hidden" name="txtOpc" value="<%=str_Opc%>">
<%if request("excel")<>1 then%>              
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
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
          <td width="26"></td>
          <td width="50"></td>
          <td width="26"></td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font><a href="rel_transmega.asp?selMegaProcesso=<%=mega%>&amp;excel=1" target="_blank"><img border="0" src="../imagens/exp_excel.gif"></a></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <%end if%>
              <p><font face="Verdana" size="3" color="#330099">Relatório de
              Transações x Mega-Processo</font></p>
              <%set rsmega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)%>
              <p><b><font face="Verdana" color="#330099" size="2">Mega-Processo
              : <%=mega%> - <%=rsmega("MEPR_TX_DESC_MEGA_PROCESSO")%></font></b></p>
              <table border="0" width="98%" cellspacing="3" cellpadding="2">
              <tr>
                  <td width="56%" bgcolor="#330099"><b><font face="Verdana" color="#FFFFFF" size="1">Transação</font></b></td>
                  <td width="63%" bgcolor="#330099"><b><font face="Verdana" color="#FFFFFF" size="1">Mega-Processo</font></b></td>
                  <td width="25%" bgcolor="#330099"><b><font face="Verdana" color="#FFFFFF" size="1">Chave
                    do Cadastrador</font></b></td>
                </tr>
	          <%
	          tem=0
              DO UNTIL RS.EOF=TRUE
              SET RSTEMP=DB.EXECUTE("SELECT DISTINCT TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO, ATUA_CD_NR_USUARIO FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE TRAN_CD_TRANSACAO='" & RS("TRAN_CD_TRANSACAO") & "' AND MEPR_CD_MEGA_PROCESSO <>" & mega & " ORDER BY TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO")

              IF RSTEMP.EOF=FALSE THEN
              
              SET RSTEMP2=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & RS("TRAN_CD_TRANSACAO") & "'")
              
              VALOR=RS("TRAN_CD_TRANSACAO") & "-" & RSTEMP2("TRAN_TX_DESC_TRANSACAO")
              COR="#C4C8A4"
              DO UNTIL RSTEMP.EOF=TRUE
                %>
                <tr>
                  <td width="56%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=VALOR%></font></td>
                  <%set rsmega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RSTEMP("MEPR_CD_MEGA_PROCESSO"))%>
                  <td width="63%" bgcolor="#9CCFCF"><font face="Verdana" size="1"><%=RSMEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
                  <td width="25%" bgcolor="#9CCFCF"><font face="Verdana" size="1"><%=RSTEMP("ATUA_CD_NR_USUARIO")%></font></td>
                </tr>
              <%
              tem=tem+1
              RSTEMP.MOVENEXT
              VALOR=""
              COR="WHITE"
              LOOP
              END IF
              RS.MOVENEXT
              LOOP
              %>
              </table>
<%if tem=0 then%>              
  <p><font face="Verdana" size="2" color="#800000"><b>&nbsp;Nenhum Registro
  Encontrado para a Seleção</b></font></p>
  <%end if%>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")

ssql="SELECT * FROM " & Session("PREFIXO") & "TRANSACAO_MEGA WHERE MEPR_CD_MEGA_PROCESSO=" & mega & " ORDER BY TRAN_CD_TRANSACAO"

set rs=db.execute(ssql)
%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="gera_rel_megaemp.asp">
              <input type="hidden" name="txtEmpSelecionada"><input type="hidden" name="txtOpc" value="<%=str_Opc%>">
<%if request("excel")<>1 then%>              
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
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
          <td width="26"></td>
          <td width="50"></td>
          <td width="26"></td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font><a href="rel_transmega.asp?selMegaProcesso=<%=mega%>&amp;excel=1" target="_blank"><img border="0" src="../imagens/exp_excel.gif"></a></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <%end if%>
              <p><font face="Verdana" size="3" color="#330099">Relatório de
              Transações x Mega-Processo</font></p>
              <%set rsmega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)%>
              <p><b><font face="Verdana" color="#330099" size="2">Mega-Processo
              : <%=mega%> - <%=rsmega("MEPR_TX_DESC_MEGA_PROCESSO")%></font></b></p>
              <table border="0" width="98%" cellspacing="3" cellpadding="2">
              <tr>
                  <td width="56%" bgcolor="#330099"><b><font face="Verdana" color="#FFFFFF" size="1">Transação</font></b></td>
                  <td width="63%" bgcolor="#330099"><b><font face="Verdana" color="#FFFFFF" size="1">Mega-Processo</font></b></td>
                  <td width="25%" bgcolor="#330099"><b><font face="Verdana" color="#FFFFFF" size="1">Chave
                    do Cadastrador</font></b></td>
                </tr>
	          <%
	          tem=0
              DO UNTIL RS.EOF=TRUE
              SET RSTEMP=DB.EXECUTE("SELECT DISTINCT TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO, ATUA_CD_NR_USUARIO FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE TRAN_CD_TRANSACAO='" & RS("TRAN_CD_TRANSACAO") & "' AND MEPR_CD_MEGA_PROCESSO <>" & mega & " ORDER BY TRAN_CD_TRANSACAO, MEPR_CD_MEGA_PROCESSO")

              IF RSTEMP.EOF=FALSE THEN
              
              SET RSTEMP2=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & RS("TRAN_CD_TRANSACAO") & "'")
              
              VALOR=RS("TRAN_CD_TRANSACAO") & "-" & RSTEMP2("TRAN_TX_DESC_TRANSACAO")
              COR="#C4C8A4"
              DO UNTIL RSTEMP.EOF=TRUE
                %>
                <tr>
                  <td width="56%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=VALOR%></font></td>
                  <%set rsmega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RSTEMP("MEPR_CD_MEGA_PROCESSO"))%>
                  <td width="63%" bgcolor="#9CCFCF"><font face="Verdana" size="1"><%=RSMEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
                  <td width="25%" bgcolor="#9CCFCF"><font face="Verdana" size="1"><%=RSTEMP("ATUA_CD_NR_USUARIO")%></font></td>
                </tr>
              <%
              tem=tem+1
              RSTEMP.MOVENEXT
              VALOR=""
              COR="WHITE"
              LOOP
              END IF
              RS.MOVENEXT
              LOOP
              %>
              </table>
<%if tem=0 then%>              
  <p><font face="Verdana" size="2" color="#800000"><b>&nbsp;Nenhum Registro
  Encontrado para a Seleção</b></font></p>
  <%end if%>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
