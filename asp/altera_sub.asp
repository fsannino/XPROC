<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

'if isnull(session("MegaProcesso")) then
'	str_mega = Request("selMegaProcesso")
'else
'	str_mega=session("MegaProcesso")
'end if

str_proc=request("selProcesso")

on error resume next
str_SQL = ""
str_SQL = str_SQL & " SELECT * "
str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL = str_SQL & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL = str_SQL & " ORDER BY MEPR_TX_DESC_MEGA_PROCESSO "

set rs1=db.execute(str_SQL)

if str_mega <> 0 then
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega )
else
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

if str_mega=0 and str_proc=0 then
set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0 ORDER BY SUPR_TX_DESC_SUB_PROCESSO")
else
set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc & " ORDER BY SUPR_TX_DESC_SUB_PROCESSO")
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
<form method="POST" action="altera_sub.asp">
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
      <td width="50%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Alteração 
        de Sub-Processos</font></td>
      <td width="26%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="0%"></td>
      <td width="5%"></td>
      <td width="94%"></td>
      <td width="1%"></td>
    </tr>
    <tr>
      <td width="0%"></td>
      <td width="5%">&nbsp;</td>
      <td width="94%">&nbsp;</td>
      <td width="1%"></td>
    </tr>
    <tr> 
      <td width="0%"></td>
      <td width="5%">&nbsp;</td>
      <td width="94%">&nbsp;</td>
      <td width="1%"></td>
    </tr>
    <tr> 
      <td width="0%"></td>
      <td width="5%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Mega 
        Processo</b></font></td>
      <td width="94%"> 
        <select size="1" name="selMegaProcesso" onchange="javascript:submit()">
          <option value="0">Selecione o Mega Processo</option>
          <%
       do until rs1.eof=true
       if trim(str_mega)=trim(rs1("MEPR_CD_MEGA_PROCESSO")) then
       %>
          <option selected value=<%=rs1("MEPR_CD_MEGA_PROCESSO")%>><%=rs1("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
       else
       %>
          <option value=<%=rs1("MEPR_CD_MEGA_PROCESSO")%>><%=rs1("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
       end if
       rs1.movenext
       loop
       %>
        </select>
      </td>
      <td width="1%"></td>
    </tr>
    <tr> 
      <td width="0%"></td>
      <td width="5%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Processo</b></font></td>
      <td width="94%"> 
        <select size="1" name="selProcesso" onchange="javascript:submit()">
          <option value="0">Selecione o Processo</option>
          <%
       do until rs2.eof=true
       if trim(str_proc)=trim(rs2("PROC_CD_PROCESSO")) then
       %>
          <option selected value=<%=rs2("PROC_CD_PROCESSO")%>><%=rs2("PROC_TX_DESC_PROCESSO")%></option>
          <%
       else
       %>
          <option value=<%=rs2("PROC_CD_PROCESSO")%>><%=rs2("PROC_TX_DESC_PROCESSO")%></option>
          <%
       end if
       rs2.movenext
       loop
       %>
        </select>
      </td>
      <td width="1%"></td>
    </tr>
    <%
if str_Mega>0 and str_proc>0 then 
set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc)
%>
    <tr> 
      <td width="0%"></td>
      <td width="5%"></td>
      <td width="94%"></td>
      <td width="1%"></td>
    </tr>
    <tr> 
      <td width="0%"></td>
      <td width="5%"></td>
      <td width="94%"></td>
      <td width="1%"></td>
    </tr>
    <tr> 
      <td width="0%"></td>
      <td width="5%"></td>
      <td width="94%"></td>
      <td width="1%"></td>
    </tr>
    <%end if%>
    <tr> 
      <td width="0%">&nbsp;</td>
      <td width="5%">&nbsp;</td>
      <%if rs.eof=true and str_proc<>0 then%>
      <td width="94%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#800000">Nenhum 
        Registro Encontrado</font></b></td>
      <%end if%>
      <td width="1%">&nbsp;</td>
    </tr>
    <%if rs.eof=false then%>
    <tr> 
      <td width="0%">&nbsp;</td>
      <td width="5%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></b></td>
      <td width="94%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Sub-Processo</font></b></td>
      <td width="1%">&nbsp;</td>
    </tr>
    <%end if%>
    <tr> 
      <td width="0%">&nbsp;</td>
      <td width="5%">&nbsp;</td>
      <td width="94%">&nbsp;</td>
      <td width="1%">&nbsp;</td>
    </tr>
    <%do while not rs.EOF %>
    <tr> 
      <td width="0%">&nbsp;</td>
      <td width="5%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUPR_CD_SUB_PROCESSO")%></font></td>
      <td width="94%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="altera_sub1.asp?selMegaProcesso=<%=str_Mega%>&selProcesso=<%=rs("PROC_CD_PROCESSO")%>&selSubProcesso=<%=rs("SUPR_CD_SUB_PROCESSO")%>"><%=rs("SUPR_TX_DESC_SUB_PROCESSO")%></font></td>
      <td width="1%">&nbsp;</td>
    </tr>
    <% rs.movenext
  Loop
  %>
    <tr> 
      <td width="0%">&nbsp;</td>
      <td width="5%">&nbsp;</td>
      <td width="94%">&nbsp;</td>
      <td width="1%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="0%">&nbsp;</td>
      <td width="5%">&nbsp;</td>
      <td width="94%">&nbsp;</td>
      <td width="1%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="0%">&nbsp;</td>
      <td width="5%">&nbsp;</td>
      <td width="94%">&nbsp;</td>
      <td width="1%">&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<p>&nbsp;</p><!-- #EndEditable -->
</body>
</html>
