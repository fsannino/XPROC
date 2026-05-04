<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
str_mega=request("mega")
str_proc=request("proc")
str_sub=request("sub")

select case request("ORDER")
	CASE 1
		VALOR="EMPR_CD_NR_EMPRESA "
	CASE 2
		VALOR="EMPR_TX_NOME_EMPRESA "
	CASE ELSE
		VALOR="EMPR_TX_NOME_EMPRESA "
END SELECT


set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

if str_sub=0 then
	set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE ORDER BY " & valor)
else
	set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc & " AND SUPR_CD_SUB_PROCESSO=" & str_sub)
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
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
      das Empresas Cadastradas</font></td>
    <td width="26%">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <%if str_sub=0 then%>
  <tr> 
    <td width="9%"></td>
    <td width="13%"></td>
    <td width="62%"><font size="1" face="Verdana"><b>Clique na coluna desejada
      para ordenar</b></font></td>
    <td width="16%"></td>
  </tr>
  <%END IF%>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <%
    if str_sub>0 then
    set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc & " AND SUPR_CD_SUB_PROCESSO=" & str_sub)
    valor=rs1("SUPR_TX_DESC_SUB_PROCESSO")
    %>
    <td width="62%">&nbsp;<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080"><b>Empresas 
      para o Sub-Processo : </b><%=str_sub%> - <%=valor%></font></td>
    <%end if%>
    <td width="16%">&nbsp;</td>
  </tr>
  <%if str_ativ=0 then%>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></b></td>
    <td width="62%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Empresas</font></b></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%ELSE%>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;digo</font></b></td>
    <td width="62%" bgcolor="#0066CC"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Empresas</font></b></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%END IF%>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%
  do while not rs.EOF
  set rs_=db.execute("SELECT * FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE WHERE EMPR_CD_NR_EMPRESA=" & rs("EMPR_CD_NR_EMPRESA")) 	  
  nome_empresa= RS_("EMPR_TX_NOME_EMPRESA")
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("EMPR_CD_NR_EMPRESA")%></font></td>
    <td width="62%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=nome_empresa%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%
  rs.movenext
  Loop
  %>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="9%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p> 
</body>
</html>
