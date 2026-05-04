<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega=request("selMegaProcesso")
str_proc=request("selProcesso")

ordena=request("order")

select case ordena
	case 1
		valor="SUPR_CD_SUB_PROCESSO"
	case 2
		valor="SUPR_TX_DESC_SUB_PROCESSO"
	case 3
		valor="SUPR_NR_SEQUENCIA"
	case else
		valor="SUPR_TX_DESC_SUB_PROCESSO"
end select

on error resume next

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

if str_mega <> 0 then
	set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega )
else
	set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

if str_mega=0 and str_proc=0 then
	set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0 ORDER BY " & valor)
	str_mega=0
	str_proc=0
else
	set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc & " ORDER BY " & valor)
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="POST" action="consulta_sub.asp">
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
      <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
        dos Sub-Processos Cadastrados</font> 
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
      </td>
      <td width="26%">&nbsp;</td>
    </tr>
  </table>
  <table width="775" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="20"></td>
      <td width="114"></td>
      <td colspan="2" width="441"></td>
      <td width="37"></td>
    </tr>
    <tr> 
      <td width="20"></td>
      <td width="114"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Mega-Processo</b></font></td>
      <td colspan="2" width="441"> 
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
      <td width="37"></td>
    </tr>
    <tr> 
      <td width="20"></td>
      <td width="114"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Processo</b></font></td>
      <td colspan="2" width="441"> 
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
      <td width="37"></td>
    </tr>
    <%
if str_Mega>0 and str_proc>0 then 
set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc)
%>
    <tr> 
      <td width="20"></td>
      <td width="114"></td>
      <td colspan="2" width="441"></td>
      <td width="37"></td>
    </tr>
    <tr> 
      <td width="20"></td>
      <td width="114"></td>
      <td colspan="2" width="441"></td>
      <td width="37"></td>
    </tr>
    <tr> 
      <td width="20"></td>
      <td width="114"></td>
    <%if rs.eof=false then%>
	  <td colspan="2" width="441"><b><font size="1" face="Verdana">Clique na coluna desejada
        para ordenar</font></b></td><%end if%>
    </tr>
    <%end if%>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="114">&nbsp;</td>
      <%if rs.eof=true and str_proc<>0 then%>
      <td colspan="2" width="441"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#800000">Nenhum 
        Registro Encontrado</font></b></td>
      <%end if%>
      <td width="37">&nbsp;</td>
    </tr>
    <%if rs.eof=false then%>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="114" bgcolor="#0066CC"><b><a href="consulta_sub.asp?order=1&selMegaProcesso=<%=str_mega%>&selProcesso=<%=str_proc%>"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></a></b></td>
      <td width="290" bgcolor="#0066CC"><b><a href="consulta_sub.asp?order=2&selMegaProcesso=<%=str_mega%>&selProcesso=<%=str_proc%>"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Sub-Processo</font></a></b></td>
      <td width="149" bgcolor="#0066CC"><b><a href="consulta_sub.asp?order=3&selMegaProcesso=<%=str_mega%>&selProcesso=<%=str_proc%>"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Sequência</font></a></b></td>
      <td width="37" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="153" bgcolor="#FFFFFF">&nbsp;</td>
    </tr>
    <%end if%>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="114">&nbsp;</td>
      <td width="290">&nbsp;</td>
      <td width="149"></td>
      <td width="37">
        <p align="center"></td>
      <td width="153">&nbsp;</td>
    </tr>
    <%do while not rs.EOF %>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="114"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUPR_CD_SUB_PROCESSO")%></font></td>
      <td width="290"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUPR_TX_DESC_SUB_PROCESSO")%></font></td>
      <td width="149"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("SUPR_NR_SEQUENCIA")%></font></td>
      <td width="37"><a href="consulta_empresa2.asp?mega=<%=str_mega%>&proc=<%=str_proc%>&sub=<%=rs("SUPR_CD_SUB_PROCESSO")%>"><img border="0" src="../imagens/icon_empresa.gif" alt="Relação de Empresas"></a></td>
      <td width="153">&nbsp;</td>
    </tr>
    <% rs.movenext
  Loop
  %>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="114">&nbsp;</td>
      <td width="290">&nbsp;</td>
      <td width="149"></td>
      <td width="37"></td>
      <td width="153">&nbsp;</td>
    </tr>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="114">&nbsp;</td>
      <td width="290">&nbsp;</td>
      <td width="149"></td>
      <td width="37"></td>
      <td width="153">&nbsp;</td>
    </tr>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="114">&nbsp;</td>
      <td width="290">&nbsp;</td>
      <td width="149"></td>
      <td width="37"></td>
      <td width="153">&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
