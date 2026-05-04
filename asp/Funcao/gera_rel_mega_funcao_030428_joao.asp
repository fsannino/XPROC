 
<%
server.scripttimeout=99999999

if request("opt") = 1 then
   Response.Buffer = TRUE
   Response.ContentType = "application/vnd.ms-excel"
end if

on error resume next

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")
func=request("selFuncao")

set temp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='"& func &"'")
func=temp("FUNE_CD_FUNCAO_NEGOCIO_PAI")
assunto=temp("SUMO_NR_SEQUENCIA")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE FUNE_CD_FUNCAO_NEGOCIO='"& func &"'")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0" link="#000000" vlink="#000000" alink="#000000">
<% if request("opt") <> 1 then %>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
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
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
          <td width="50"><a href="javascript:print()"><img border="0" src="../../imagens/print.gif"></a></td>
          <td width="26">&nbsp;</td>
          <td width="195"> 
            <p align="center"><a href="gera_rel_mega_funcao.asp?selMegaProcesso=<%=mega%>&amp;selFuncao=<%=func%>&amp;opt=1" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a>
          </td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<% end if %>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<p style="margin-top: 0; margin-bottom: 0">

&nbsp;&nbsp;&nbsp;<font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">
Relatório Fun&ccedil;&atilde;o R/3 x Transa&ccedil;&atilde;o</font> 
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)%>
<p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="2"><b>&nbsp;&nbsp;&nbsp;
Mega-Processo
: </b><%=mega%> - <%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "SUB_MODULO WHERE MEPR_CD_MEGA_PROCESSO=" & mega & " AND SUMO_NR_SEQUENCIA=" & assunto  )%>
<p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="2"><b>&nbsp;&nbsp;&nbsp; 
  Assunto : </b><%=TEMP("SUMO_TX_DESC_SUB_MODULO")%>&nbsp;</font></p>
<%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & func & "'")%>
<p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="2"><b>&nbsp;&nbsp;&nbsp; 
  Fun&ccedil;&atilde;o R/3 : </b><%=FUNC%> - <%=TEMP("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="2"><b>&nbsp;&nbsp;&nbsp; 
  Descri&ccedil;&atilde;o da Fun&ccedil;&atilde;o : </b><%=TEMP("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></font></p>
<%IF RS.EOF=FALSE THEN%>
<table border="0" cellspacing="1" cellpadding="2" width="777" bordercolor="#000000">
  <tr> 
    <td width="119" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Mega-Processo</font></b></td>
    <td width="161" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Processo</font></b></td>
    <td width="134" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Sub-Processo</font></b></td>
    <td width="175" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Atividade</font></b></td>
    <td width="146" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Transação</font></b></td>
  </tr>
  <%DO UNTIL RS.EOF=TRUE%>
  <tr> 
    <td width="119"> 
      <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO"))%>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></font> 
    </td>
    <td width="161"> 
      <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO=" & RS("PROC_CD_PROCESSO"))%>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=TEMP("PROC_TX_DESC_PROCESSO")%></font> 
    </td>
    <td width="134"> 
      <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO=" & RS("PROC_CD_PROCESSO") & " AND SUPR_CD_SUB_PROCESSO=" & RS("SUPR_CD_SUB_PROCESSO"))%>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=TEMP("SUPR_TX_DESC_SUB_PROCESSO")%></font> 
    </td>
    <td width="175"> 
      <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & RS("ATCA_CD_ATIVIDADE_CARGA"))%>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=TEMP("ATCA_TX_DESC_ATIVIDADE")%></font> 
    </td>
    <td width="146"> 
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=RS("TRAN_CD_TRANSACAO")%></font> 
    </td>
  </tr>
  <tr> 
    <td colspan="5" width="767"> 
      <div align="right"> 
        <%SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & RS("TRAN_CD_TRANSACAO") & "'") %>
        <font face="Verdana" size="1"><%=TEMP("TRAN_TX_DESC_TRANSACAO")%></font> </div>
    </td>
  </tr>
  <%
  RS.MOVENEXT
  LOOP
  %>
</table>
<%ELSE%>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;<b><font face="Verdana" size="2" color="#800000">
Nenhum
Registro Encontrado</font></b></p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<%END IF%>

</body>

</html>