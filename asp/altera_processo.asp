<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%

str_MegaProcesso = Request("selMegaProcesso")

if str_MegaProcesso=0 and not isnull(session("MegaProcesso")) then
	str_MegaProcesso=session("MegaProcesso")
end if

set db = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.recordset")

db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQL = str_SQL & " SELECT * "
str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL = str_SQL & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL = str_SQL & " ORDER BY MEPR_TX_DESC_MEGA_PROCESSO "

set rs_mega=db.execute(str_SQL)

if len(str_MegaProcesso)>0 then
   set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " ORDER BY PROC_TX_DESC_PROCESSO")
ELSE
   set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = 0 ORDER BY PROC_TX_DESC_PROCESSO")
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
<script>
function manda()
{
if(document.frm1.selProcesso.value!=0)
{
window.location.href='altera_processo1.asp?selMegaProcesso='+ document.frm1.selMegaProcesso.value +'&selProcesso='+ document.frm1.selProcesso.value
}
}
</script>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="POST" action="altera_processo.asp">
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
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" border="0"></a></div>
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
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Alteração 
        de Processos&nbsp;</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Mega-Processo</b></font></td>
      <td width="59%"> 
        <select size="1" name="selMegaProcesso" onchange="javascript:submit()">
          <option value="0">Selecione o Mega-Processo</option>
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
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <%if str_megaprocesso>0 then
    set rs_=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_megaprocesso)
    valor=rs_("MEPR_TX_DESC_MEGA_PROCESSO")
    %>
      <%END IF%>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <%if rs.eof=true and str_MegaProcesso <> 0 then%>
      <td width="59%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#800000">Nenhum 
        Registro Encontrado</font></b></td>
      <%end if%>
      <td width="14%">&nbsp;</td>
    </tr>
    <%IF RS.EOF=FALSE THEN%>
    <%end if%>
    <tr> 
      <td width="9%">&nbsp;</td>
      <%if rs.eof=false then%>
      <td width="18%">&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Processo</b></font></td>
      <td width="59%"> 
        <select size="1" name="selProcesso" onchange="javascript:manda()">
          <option value="0">Selecione o Processo</option>
          <%
        do until rs.eof=true
        %>
          <option value="<%=rs("PROC_CD_PROCESSO")%>"><%=rs("PROC_TX_DESC_PROCESSO")%></option>
          <%
        rs.movenext
        loop
        %>
        </select>
      </td>
      <%end if%>
      <td width="14%">&nbsp;</td>
    </tr>
    <%do until rs.eof=true %>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("PROC_CD_PROCESSO")%></font></td>
      <td width="59%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="altera_processo1.asp?selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=rs("PROC_CD_PROCESSO")%>"><%=rs("PROC_TX_DESC_PROCESSO")%></a></font></td>
      <td width="14%">&nbsp;</td>
    </tr>
    <% rs.movenext
  Loop
  %>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
