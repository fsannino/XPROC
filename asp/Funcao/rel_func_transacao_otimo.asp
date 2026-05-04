 
<%
server.scripttimeout=99999999

'mega1=request("selMegaFuncao")

'valor=request("selMegaProcesso")

mega1=4
valor=4

proc=request("selProcesso")
subproc=request("selSubProcesso")

tatual=0

if proc<>0 then
	complemento=" AND PROC_CD_PROCESSO=" & proc
end if

if subproc<>0 then
	complemento=complemento+" AND SUPR_CD_SUB_PROCESSO=" & subproc
end if

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO, FUNE_TX_TITULO_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & mega1 & " ORDER BY FUNE_CD_FUNCAO_NEGOCIO")
set temp=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega1)

texto=temp("MEPR_TX_DESC_MEGA_PROCESSO")
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
          <td width="195"></td>
            <td width="27"></td>  <td width="50"><a href="rel_func_transacao_excel.asp?selMegaFuncao=<%=mega1%>&selMegaProcesso=<%=valor%>&selProcesso=<%=proc%>&selSubProcesso=<%=subproc%>" target="blank"><img border="0" src="../../imagens/exp_excel.gif"></a></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório de Associação de
Funções de Negócio</font></p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#330099" size="2">Mega-Processo selecionado :
<%=valor%>-<%=texto%></font></b></p>
<p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp;

</font>

<table border="1" cellspacing="0" cellpadding="0" height="43">
  <tr>
    <td width="350" bgcolor="#330099" colspan="2" height="13">
      <p align="right" style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1" color="#FFFFFF">Função --&gt;</font></b></p>
    </td>
	<%do until rs.eof=true%>      
    <td width="20" rowspan="2" height="22" bgcolor="#E1E1E1"><font face="Verdana" color="#330099" size="1"><%=rs("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
    <%
    rs.movenext
    loop
    %>
  </tr>
  <tr>
    <td bgcolor="#330099" height="7">
      <p align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Mega-Processo</font></b></td>
    <td width="150" align="right" bgcolor="#330099" height="7">
      <p align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Transação</font></b></td>
  </tr>
  <%
  rs.movefirst
  
  valor_sql="SELECT DISTINCT MEPR_CD_MEGA_PROCESSO, TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & valor & complemento & " ORDER BY MEPR_CD_MEGA_PROCESSO, TRAN_CD_TRANSACAO"
  set rs1=db.execute(valor_sql)
  
  do until rs1.eof=true
  mega_atual=rs1("MEPR_CD_MEGA_PROCESSO")
  
  if cor="#E1E1E1" then
 			cor="#DBDED1"
 		else
 			cor="#E1E1E1"
 		end if
  %>
 <tr> 
 <td width="200" height="10" bgcolor="<%=cor%>">
 		<%
 		if mega_ant<>mega_atual then
	 		set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual)
 			valor_mega=rstemp("MEPR_TX_DESC_MEGA_PROCESSO")
 		else
 			valor_mega="-"
 		end if
  		%>
       <p align="center"><font face="Verdana" size="1" color="#330099"><%=valor_mega%></font></td>
    <td width="150" height="10" bgcolor="<%=cor%>">
      <p align="center"><font face="Verdana" size="1" color="#330099"><%=RS1("tran_cd_transacao")%></font></td>
	 
	 <%
	 do until rs.eof=true
	 
	 set tem=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RS1("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "' AND TRAN_CD_TRANSACAO='" & RS1("TRAN_CD_TRANSACAO") & "'")	 
	 
	 if tem.eof=false then
	 %>
	 <td width="27" height="10" bgcolor="<%=cor%>">
      <p align="center"><font face="AdLib BT" color="#330099" size="4">X</font></p>
 </td>
		<%else%>    
    <td width="161" height="10" bgcolor="<%=cor%>"></td>
    <%end if
    rs.movenext
    loop
    
    rs.movefirst
    
    %>
    </tr>
   <%
   mega_ant=rs1("MEPR_CD_MEGA_PROCESSO")
   rs1.movenext
   loop
   %>
</table>
</body>
</html>
