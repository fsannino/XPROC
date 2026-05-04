<%
if request("excel")=1 then
	'Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

dim status

mega=request("selMegaProcesso")
proc=request("selProcesso")
subp=request("selSubProcesso")
onda=request("selOnda")
status = request("selStatus")
str_Escopo = request("selEscopo")

cenario1=request("ID")
cenario2=request("ID2")
str_Assunto=0
str_Assunto=request("selAssunto")

if cenario1="0" and cenario2="0" then
   if mega<>0 then
	  compl=compl+"MEPR_CD_MEGA_PROCESSO=" & mega & " AND "
   end if
   if proc<>0 then
	  compl=compl+"PROC_CD_PROCESSO=" & proc & " AND "
   end if
   if subp<>0 then
	compl=compl+"SUPR_CD_SUB_PROCESSO=" & subp & " AND "
	end if
	if onda<>0 then
		compl=compl+"ONDA_CD_ONDA=" & onda & " AND "
	ELSE
		compl=compl+"ONDA_CD_ONDA<>4 AND "
	end if
	if status<>"0" then
		compl=compl+"CENA_TX_SITUACAO='" & status & "' AND "
	end if
	if str_Assunto<>"0" then
		compl=compl+" MEPR_CD_MEGA_PROCESSO_SUMO =" & mega & " AND SUMO_NR_SEQUENCIA='" & str_Assunto & "' AND "
	end if
	if str_Escopo <> 2 then
		compl=compl+"CENA_TX_SITUACAO_VALIDACAO='" & str_Escopo & "' AND "
	end if
	
	tamanho=len(compl)
	tamanho=tamanho-5
	compl=left(compl,tamanho)
else
	if cenario1<>"0" then
		compl="CENA_CD_CENARIO='" & cenario1& "'"
		cenario=cenario1
	else
		if cenario2<>"0" then
			compl="CENA_CD_CENARIO='" & cenario2& "'"
			cenario=cenario2
		end if
	end if
end if

if len(compl)>0 then
	compl=" WHERE " & compl
end if

ordem=request("ORDER")

if ordem="" then
	ordem="CENA_CD_CENARIO"
end if

ordem=" ORDER BY " & ordem

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO" & compl & ordem

SSQL1=SSQL

if request("excel")=1 then
	ssql=request("ssql2")
end if

'response.Write(ssql)

set rs=db.execute(ssql)

IF RS.EOF=TRUE THEN
	TEM=0
ELSE
	TEM=1
END IF
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

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%if request("excel")<>1 then%>
<form name="frm1" method="POST" action="">
  <table width="773" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td height="66" colspan="2">&nbsp;</td>
      <td height="66" colspan="2">&nbsp;</td>
      <td valign="top" colspan="2" height="66"> 
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="111">&nbsp; </td>
      <td height="20" width="30">&nbsp;</td>
      <td height="20" width="213"><a href="gera_rel_status.asp?ssql2=<%=ssql1%>&amp;excel=1" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a></td>
      <td colspan="2" height="20">&nbsp;
        
      </td>
      <td height="20" width="334">&nbsp;</td>
    </tr>
  </table>
  <%end if%>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  <p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="3">Relatório
  Cenário x Status</font> </p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  <table border="0" width="97%" cellspacing="0" cellpadding="0">
    <%if tem=1 then%>
    <tr>
      <td width="23%" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Cenário</b></font></td>
      <td width="18%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Escopo</b></font></td>
      <td width="15%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Status</b></font></td>
      <td width="19%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Tipo</b></font></td>
      <td width="18%" bgcolor="#330099" align="center"><font face="Verdana" size="2" color="#FFFFFF"><b>Configuração</b></font></td>
      <td width="21%" bgcolor="#330099" align="center">
        <p align="left"><font face="Verdana" size="2" color="#FFFFFF"><b>Desenvolvimento</b></font></p>
      </td>
    </tr>
    <%end if%>
    <%do until rs.eof=true
    IF COR="#E4E4E4" THEN
    	COR="WHITE"
    ELSE
    	COR="#E4E4E4"
    END IF
    %>
    <tr>
      <td width="23%" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=rs("CENA_CD_CENARIO")%>-<%=RS("CENA_TX_TITULO_CENARIO")%></font></td>
      <%
      escopo = rs("CENA_TX_SITUACAO_VALIDACAO")
      
      select case escopo
      	case 0
      		val_escopo="FORA DO ESCOPO"
      	case 1
      		val_escopo="DENTRO DO ESCOPO"
		end SELECT
      
      
      If rs("CENA_TX_SITUACAO") = "DF" Then
			      ls_Situacao_Cenario = "DEFINIDO"
			   elseIf rs("CENA_TX_SITUACAO") = "EE" Then
			      ls_Situacao_Cenario = "EM ELABORAÇÃO"
		      elseIf rs("CENA_TX_SITUACAO") = "DS" Then
				      ls_Situacao_Cenario = "DESENHADO"
			   elseIf rs("CENA_TX_SITUACAO") = "PT" Then
				      ls_Situacao_Cenario = "PRONTO PARA TESTE"
				elseIf rs("CENA_TX_SITUACAO") = "TD" Then
				      ls_Situacao_Cenario = "TESTADO NO PED"
				elseIf rs("CENA_TX_SITUACAO") = "TQ" Then
				      ls_Situacao_Cenario = "TESTADO NO PEQ"
			   end if
      %>
      <td width="18%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=VAL_ESCOPO%></font></td>
      <td width="15%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=ls_Situacao_Cenario%></font></td>
      <%
      if rs("CENA_TX_SITU_DESENHO_TIPO")=1 THEN
      		SITUACAO="COM DESENVOLVIMENTO"
      ELSE
      IF rs("CENA_TX_SITU_DESENHO_TIPO")=2 THEN
      		SITUACAO="SEM DESENVOLVIMENTO"
      ELSE
      		SITUACAO="  "
      END IF
      END IF
      %>
      <td width="19%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO%></font></td>
      <%
      if rs("CENA_TX_SITU_DESENHO_CONF")=0 THEN
      		SITUACAO2="  "
      ELSE
      IF rs("CENA_TX_SITU_DESENHO_CONF")=1 THEN
      		SITUACAO2="CONFIGURAÇÃO CONCLUÍDA"
      ELSE
      		SITUACAO2="  "
      END IF
      END IF
      %>
      <td width="18%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO2%></font></td>
      <%
      if rs("CENA_TX_SITU_DESENHO_DESE")=0 THEN
      		SITUACAO3="  "
      ELSE
      IF rs("CENA_TX_SITU_DESENHO_DESE")=1 THEN
      		SITUACAO3="DESENVOLVIMENTO CONCLUÍDO"
      ELSE
      		SITUACAO3="  "
      END IF
      END IF
      %>
      <td width="21%" align="center" bgcolor="<%=COR%>"><font face="Verdana" size="1"><%=SITUACAO3%></font></td>
    </tr>
  
    <%
      rs.movenext
      loop
      %>
</table>
<%if tem=0 then%>
  <font color="#800000" face="Verdana" size="2"><b>Nenhum Registro Encontrado</b></font>
 <%end if%>
  </form>
<p></p>
</body>
</html>
