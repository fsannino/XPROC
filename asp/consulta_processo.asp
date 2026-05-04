<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
str_MegaProcesso = Request("selMegaProcesso")

ordena=request("order")

select case ordena
	case 1
		valor="PROC_CD_PROCESSO"
	case 2
		valor="PROC_TX_DESC_PROCESSO"
	case 3
		valor="PROC_NR_SEQUENCIA"
	case else
		valor="PROC_TX_DESC_PROCESSO"
end select

set db = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.recordset")

db.Open Session("Conn_String_Cogest_Gravacao")

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

if len(str_MegaProcesso)>0 then
   set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " ORDER BY " & valor)
ELSE
   set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = 0 ORDER BY " & valor)
end if

if len(str_MegaProcesso)=0 then
	str_MegaProcesso=0
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
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" action="consulta_processo.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
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
      <td colspan="3" height="20">&nbsp; </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
        dos Processos Cadastrados</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="53%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Mega-Processo</b></font></td>
      <td width="53%">
        <select size="1" name="selMegaProcesso" onchange="javascript:submit()">
          <option value="0">Selecione o Mega Processo</option>
          <%
        DO UNTIL RS_MEGA.EOF=TRUE
        if trim(str_MegaProcesso) = trim(RS_MEGA("MEPR_CD_MEGA_PROCESSO"))then
        %>
          <option selected value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%else
        %>
          <option value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
        end if
        RS_MEGA.MOVENEXT
        LOOP
        %>
        </select>
      </td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="53%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="18%"></td>
      <%if str_megaprocesso>0 then
    set rs_=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_megaprocesso)
    valor=rs_("MEPR_TX_DESC_MEGA_PROCESSO")
    %>
      <%END IF%>
      <%if rs.eof=false then%>
      <td width="53%"><b><font size="1" face="Verdana">Clique na
        coluna desejada para ordenar</font></b></td><%end if%>
      <td width="16%"></td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <%if rs.eof=true and str_MegaProcesso <> 0 then%>
      <td width="53%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#800000">Nenhum 
        Registro Encontrado</font></b></td>
      <%end if%>
      <td width="16%">&nbsp;</td>
    </tr>
    <%IF RS.EOF=FALSE THEN%>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%" bgcolor="#0066CC"><b><a href="consulta_processo.asp?order=1&selMegaProcesso=<%=str_MegaProcesso%>"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></a></b></td>
      <td width="41%" bgcolor="#0066CC"><font color="#FFFFFF"><b><a href="consulta_processo.asp?order=2&selMegaProcesso=<%=str_MegaProcesso%>"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Processo
        </font></a>
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
        </font></b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"> 
        (clique para ver os Sub-Processos)</font></font></td>
      <td width="14%" bgcolor="#0066CC"><b><a href="consulta_processo.asp?order=3&selMegaProcesso=<%=str_MegaProcesso%>"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Sequência</font></a></b></td>
      <td width="18%" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <%end if%>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="41%">&nbsp;</td>
      <td width="14%"></td>
      <td width="18%" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <%do until rs.eof=true %>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("PROC_CD_PROCESSO")%></font></td>
      <td width="41%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="consulta_sub.asp?selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=rs("PROC_CD_PROCESSO")%>"><%=rs("PROC_TX_DESC_PROCESSO")%></a></font></td>
      <td width="14%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("PROC_NR_SEQUENCIA")%></font></td>
      <td width="18%" bgcolor="#FFFFFF"></td>
    </tr>
    <% rs.movenext
  Loop
  %>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="41%">&nbsp;</td>
      <td width="14%"></td>
      <td width="18%" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="41%">&nbsp;</td>
      <td width="14%"></td>
      <td width="18%" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="41%">&nbsp;</td>
      <td width="14%"></td>
      <td width="18%" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
