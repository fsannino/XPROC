<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../protege/protege.asp" -->
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

ordena=request("order")
select case ordena
	case 1
		valor="" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
	case 2
		valor="" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
	case else
		valor="" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
end select

str_SQL_Fechamento = ""
str_SQL_Fechamento = str_SQL_Fechamento & " SELECT  "
str_SQL_Fechamento = str_SQL_Fechamento & " FEES_CD_FECHAMENTO, FEES_DT_FECHAMENTO, "
str_SQL_Fechamento = str_SQL_Fechamento & " FEES_DT_FECHAMENTO_ANTERIOR, "
str_SQL_Fechamento = str_SQL_Fechamento & " FEES_TX_CHAVE_QUEM_FECHOU, FEES_TX_COMENTARIO, "
str_SQL_Fechamento = str_SQL_Fechamento & " FEES_TX_CONTROLA_FECHAMENTO "
str_SQL_Fechamento = str_SQL_Fechamento & " FROM FECHA_ESCOPO "
str_SQL_Fechamento = str_SQL_Fechamento & " order by FEES_CD_FECHAMENTO "
'response.Write(str_SQL_Fechamento)
Set rdsFechamento  = Conn_db.Execute(str_SQL_Fechamento)
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
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
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="24%">&nbsp;</td>
    <td width="50%">&nbsp;</td>
    <td width="26%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="24%">&nbsp;</td>
    <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Fechamento 
      de Escopo do Projeto R/3</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td><font size="1" face="Verdana"><b>Clique na coluna desejada para ordenar</b></font></td>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="2" cellpadding="0">
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="19%"><font size="1" face="Verdana">&nbsp;</font></td>
    <td width="52%"><font size="1" face="Verdana">&nbsp;</font></td>
    <td width="15%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Identificador 
      </font></b></td>
    <td width="19%" bgcolor="#0066CC" style="color: #FFFFFF"><div align="center"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Data 
        Fechamento </font></b></div></td>
    <td width="52%" bgcolor="#0066CC" style="color: #FFFFFF"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Coment&aacute;rio</font></b></td>
    <td width="15%" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="19%">&nbsp;</td>
    <td width="52%">&nbsp;</td>
    <td width="15%">&nbsp;</td>
  </tr>
  <%do while not rdsFechamento.EOF 
       if str_Cor_Linha = "#EEEEEE" then
	      str_Cor_Linha = "#FFFFFF"
	   else
	      str_Cor_Linha = "#EEEEEE"
	   end if
  %>
    
  <tr bgcolor="<%=str_Cor_Linha%>"> 
    <td width="14%"> 
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="seleciona_historico.asp?pCodEscopo=<%=rdsFechamento("FEES_CD_FECHAMENTO")%>"><%=rdsFechamento("FEES_CD_FECHAMENTO")%></a></font></div></td>
    <td width="19%"> 
      <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsFechamento("FEES_DT_FECHAMENTO")%></font></div></td>
    <td width="52%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsFechamento("FEES_TX_COMENTARIO")%></font></td>
    <td width="15%" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <% rdsFechamento.movenext
  Loop
  %>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="19%">&nbsp;</td>
    <td width="52%">&nbsp;</td>
    <td width="15%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="19%">&nbsp;</td>
    <td width="52%">&nbsp;</td>
    <td width="15%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="19%">&nbsp;</td>
    <td width="52%">&nbsp;</td>
    <td width="15%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
<%
rdsFechamento.Close
set rdsFechamento = Nothing
conn_db.close
set conn_db = Nothing
%>
</body>
</html>
