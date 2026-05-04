<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%

Server.ScriptTimeOut=99999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

MEGA=REQUEST("selMegaProcesso")
PROC=REQUEST("selProcesso")
SUBR=REQUEST("selSubProcesso")

MODULO=REQUEST("selModulo")
ATIVIDADE=REQUEST("selAtividade")
TRANSACAO=REQUEST("selTransacao")

IF MEGA<>0 THEN
COMPL1="MEPR_CD_MEGA_PROCESSO=" & MEGA
END IF

IF PROC<>0 THEN
COMPL2="PROC_CD_PROCESSO=" & PROC
END IF

IF SUBR<>0 THEN
COMPL3="SUPR_CD_SUB_PROCESSO=" & SUBR
END IF

IF MODULO<>0 THEN
COMPL4="MODU_CD_MODULO=" & MODULO
END IF

IF ATIVIDADE<>0 THEN
COMPL5="ATCA_CD_ATIVIDADE_CARGA=" & ATIVIDADE
END IF

IF TRANSACAO<>"0" THEN
COMPL6="TRAN_CD_TRANSACAO='" & TRANSACAO & "'"
END IF

IF COMPL1<>"" THEN
COMPLE = COMPL1
END IF

IF COMPL2<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE + " AND " + COMPL2
ELSE
COMPLE=COMPL2
END IF
END IF

IF COMPL3<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL3
ELSE
COMPLE=COMPL3
END IF
END IF

IF COMPL4<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL4
ELSE
COMPLE=COMPL4
END IF
END IF

IF COMPL5<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL5
ELSE
COMPLE=COMPL5
END IF
END IF

IF COMPL6<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL6
ELSE
COMPLE=COMPL6
END IF
END IF

IF COMPLE<>"" THEN
CONECTA="WHERE "
END IF

'SSQL="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL " & CONECTA & COMPLE & " ORDER BY RELA_NR_SEQUENCIA"

SSQL="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL " & CONECTA & COMPLE & " ORDER BY MEPR_CD_MEGA_PROCESSO, PROC_CD_PROCESSO, SUPR_CD_SUB_PROCESSO, MODU_CD_MODULO, ATCA_CD_ATIVIDADE_CARGA"

'response.write ssql

set rs=db.execute(SSQL)
%>

<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0">
<form method="POST" action="gera_rel_modatca.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
          <td width="26">&nbsp;</td>
            <td width="103"><img src="../imagens/print.gif" width="90" height="35" border="0" onclick="javascript:print()"></td>
          <td width="119">
            <p align="center"><a href="gera_rel_geral_excel.asp?selMegaProcesso=<%=mega%>&selProcesso=<%=proc%>&selSubProcesso=<%=subr%>&selModulo=<%=modulo%>&selAtividade=<%=atividade%>&selTransacao=<%=transacao%>" target="_blank"><img border="0" src="../imagens/exp_excel.gif"></a></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center"></p>
  <p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" color="#330099" size="3">Relatório
  </font><font face="Verdana" color="#330099" size="3">Geral de Relações
  Definidas</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
<%if rs.eof=true then%>  
<p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" color="#800000" size="2"><b>Nenhum
Registro Encontrado</b></font></p>
<%else%>
<p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
<table border="0" width="100%">
  <tr>
    <td width="16%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Mega-Processo</font></b></td>
    <td width="16%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Processo</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Sub-Processo</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Agrupamento
      das Atividades</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Atividade</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Transação</font></b></td>
  </tr>
<%
valor1=""
valor2=""
valor3=""
valor4=""
valor5=""
valor6=""

on error resume next

mega_atual=RS("MEPR_CD_MEGA_PROCESSO")
proc_atual=RS("PROC_CD_PROCESSO")
sub_atual=RS("SUPR_CD_SUB_PROCESSO")
modulo_atual=RS("MODU_CD_MODULO")
atividade_atual=RS("ATCA_CD_ATIVIDADE_CARGA")

do until rs.eof=true

if mega_ant<>mega_atual then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual)
	valor1=rs1("MEPR_TX_DESC_MEGA_PROCESSO")
else
	valor1=""
end if

if proc_ant<>proc_atual then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual)
	valor2=rs1("PROC_TX_DESC_PROCESSO")
else
	valor2=""
end if

if sub_atual<>sub_ant then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual)
	valor3=rs1("SUPR_TX_DESC_SUB_PROCESSO")
else
	if proc_atual<>proc_ant then
		set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual)
		valor3=rs1("SUPR_TX_DESC_SUB_PROCESSO")
	else
		valor3=""
	end if
end if

if (modulo_atual<>modulo_ant) or (valor3<>"") then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & modulo_atual)
	valor4=rs1("MODU_TX_DESC_MODULO")
else
	valor4=""
end if

if(atividade_atual<>atividade_ant) or (valor4<>"") then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & Atividade_atual)
	valor5=RS1("ATCA_TX_DESC_ATIVIDADE")
else
	valor5=""
end if

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & RS("TRAN_CD_TRANSACAO") & "'")
valor6=RS("TRAN_CD_TRANSACAO") & "-" & rs1("TRAN_TX_DESC_TRANSACAO")
%>
  <tr>
    <%if valor1<>"" then%>
    <td width="16%" bgcolor="#D3D3D3"><font face="Verdana" size="1"><%=valor1%></font></td>
    <%else%>
    <td width="16%"><font face="Verdana" size="1"><%=valor1%></font></td>
    <%end if%>

    <%if valor2<>"" then%>
    <td width="16%" bgcolor="#B7C8BC"><font face="Verdana" size="1"><%=valor2%></font></td>
    <%else%>
    <td width="16%"><font face="Verdana" size="1"><%=valor2%></font></td>
    <%end if%>
    
    <%if valor3<>"" then%>
    <td width="17%" bgcolor="#FFFFCC"><font face="Verdana" size="1"><%=valor3%></font></td>
    <%else%>
    <td width="17%"><font face="Verdana" size="1"><%=valor3%></font></td>
	 <%end if%>
    
    <%if valor4<>"" then%>
    <td width="17%" bgcolor="#FF9900"><font face="Verdana" size="1"><%=VALOR4%></font></td>
    <%ELSE%>
    <td width="17%"><font face="Verdana" size="1"><%=VALOR4%></font></td>
	<%END IF%>
    
    <%if valor5<>"" then%>
    <td width="17%" bgcolor="#FFFF00"><font face="Verdana" size="1"><%=valor5%></font></td>
    <%else%>
    <td width="17%"><font face="Verdana" size="1"><%=valor5%></font></td>
    <%end if%>
    
    <td width="17%" bgcolor="#AAFFDD"><font face="Verdana" size="1"><%=valor6%></font></td>
  </tr>
<%

mega_ant=RS("MEPR_CD_MEGA_PROCESSO")
proc_ant=RS("PROC_CD_PROCESSO")
sub_ant=RS("SUPR_CD_SUB_PROCESSO")
modulo_ant=RS("MODU_CD_MODULO")
atividade_ant=RS("ATCA_CD_ATIVIDADE_CARGA")

rs.movenext

ON ERROR RESUME NEXT

mega_atual=RS("MEPR_CD_MEGA_PROCESSO")
proc_atual=RS("PROC_CD_PROCESSO")
sub_atual=RS("SUPR_CD_SUB_PROCESSO")
modulo_atual=RS("MODU_CD_MODULO")
atividade_atual=RS("ATCA_CD_ATIVIDADE_CARGA")

LOOP
end if%>
</table>
<p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
</form>

</body>

</html>
=======
 
<!--#include file="../asp/protege/protege.asp" -->
<%

Server.ScriptTimeOut=99999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

MEGA=REQUEST("selMegaProcesso")
PROC=REQUEST("selProcesso")
SUBR=REQUEST("selSubProcesso")

MODULO=REQUEST("selModulo")
ATIVIDADE=REQUEST("selAtividade")
TRANSACAO=REQUEST("selTransacao")

IF MEGA<>0 THEN
COMPL1="MEPR_CD_MEGA_PROCESSO=" & MEGA
END IF

IF PROC<>0 THEN
COMPL2="PROC_CD_PROCESSO=" & PROC
END IF

IF SUBR<>0 THEN
COMPL3="SUPR_CD_SUB_PROCESSO=" & SUBR
END IF

IF MODULO<>0 THEN
COMPL4="MODU_CD_MODULO=" & MODULO
END IF

IF ATIVIDADE<>0 THEN
COMPL5="ATCA_CD_ATIVIDADE_CARGA=" & ATIVIDADE
END IF

IF TRANSACAO<>"0" THEN
COMPL6="TRAN_CD_TRANSACAO='" & TRANSACAO & "'"
END IF

IF COMPL1<>"" THEN
COMPLE = COMPL1
END IF

IF COMPL2<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE + " AND " + COMPL2
ELSE
COMPLE=COMPL2
END IF
END IF

IF COMPL3<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL3
ELSE
COMPLE=COMPL3
END IF
END IF

IF COMPL4<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL4
ELSE
COMPLE=COMPL4
END IF
END IF

IF COMPL5<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL5
ELSE
COMPLE=COMPL5
END IF
END IF

IF COMPL6<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL6
ELSE
COMPLE=COMPL6
END IF
END IF

IF COMPLE<>"" THEN
CONECTA="WHERE "
END IF

'SSQL="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL " & CONECTA & COMPLE & " ORDER BY RELA_NR_SEQUENCIA"

SSQL="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL " & CONECTA & COMPLE & " ORDER BY MEPR_CD_MEGA_PROCESSO, PROC_CD_PROCESSO, SUPR_CD_SUB_PROCESSO, MODU_CD_MODULO, ATCA_CD_ATIVIDADE_CARGA"

'response.write ssql

set rs=db.execute(SSQL)
%>

<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0">
<form method="POST" action="gera_rel_modatca.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
          <td width="26">&nbsp;</td>
            <td width="103"><img src="../imagens/print.gif" width="90" height="35" border="0" onclick="javascript:print()"></td>
          <td width="119">
            <p align="center"><a href="gera_rel_geral_excel.asp?selMegaProcesso=<%=mega%>&selProcesso=<%=proc%>&selSubProcesso=<%=subr%>&selModulo=<%=modulo%>&selAtividade=<%=atividade%>&selTransacao=<%=transacao%>" target="_blank"><img border="0" src="../imagens/exp_excel.gif"></a></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center"></p>
  <p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" color="#330099" size="3">Relatório
  </font><font face="Verdana" color="#330099" size="3">Geral de Relações
  Definidas</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
<%if rs.eof=true then%>  
<p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" color="#800000" size="2"><b>Nenhum
Registro Encontrado</b></font></p>
<%else%>
<p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
<table border="0" width="100%">
  <tr>
    <td width="16%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Mega-Processo</font></b></td>
    <td width="16%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Processo</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Sub-Processo</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Agrupamento
      das Atividades</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Atividade</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Transação</font></b></td>
  </tr>
<%
valor1=""
valor2=""
valor3=""
valor4=""
valor5=""
valor6=""

on error resume next

mega_atual=RS("MEPR_CD_MEGA_PROCESSO")
proc_atual=RS("PROC_CD_PROCESSO")
sub_atual=RS("SUPR_CD_SUB_PROCESSO")
modulo_atual=RS("MODU_CD_MODULO")
atividade_atual=RS("ATCA_CD_ATIVIDADE_CARGA")

do until rs.eof=true

if mega_ant<>mega_atual then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual)
	valor1=rs1("MEPR_TX_DESC_MEGA_PROCESSO")
else
	valor1=""
end if

if proc_ant<>proc_atual then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual)
	valor2=rs1("PROC_TX_DESC_PROCESSO")
else
	valor2=""
end if

if sub_atual<>sub_ant then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual)
	valor3=rs1("SUPR_TX_DESC_SUB_PROCESSO")
else
	if proc_atual<>proc_ant then
		set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual)
		valor3=rs1("SUPR_TX_DESC_SUB_PROCESSO")
	else
		valor3=""
	end if
end if

if (modulo_atual<>modulo_ant) or (valor3<>"") then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & modulo_atual)
	valor4=rs1("MODU_TX_DESC_MODULO")
else
	valor4=""
end if

if(atividade_atual<>atividade_ant) or (valor4<>"") then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & Atividade_atual)
	valor5=RS1("ATCA_TX_DESC_ATIVIDADE")
else
	valor5=""
end if

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & RS("TRAN_CD_TRANSACAO") & "'")
valor6=RS("TRAN_CD_TRANSACAO") & "-" & rs1("TRAN_TX_DESC_TRANSACAO")
%>
  <tr>
    <%if valor1<>"" then%>
    <td width="16%" bgcolor="#D3D3D3"><font face="Verdana" size="1"><%=valor1%></font></td>
    <%else%>
    <td width="16%"><font face="Verdana" size="1"><%=valor1%></font></td>
    <%end if%>

    <%if valor2<>"" then%>
    <td width="16%" bgcolor="#B7C8BC"><font face="Verdana" size="1"><%=valor2%></font></td>
    <%else%>
    <td width="16%"><font face="Verdana" size="1"><%=valor2%></font></td>
    <%end if%>
    
    <%if valor3<>"" then%>
    <td width="17%" bgcolor="#FFFFCC"><font face="Verdana" size="1"><%=valor3%></font></td>
    <%else%>
    <td width="17%"><font face="Verdana" size="1"><%=valor3%></font></td>
	 <%end if%>
    
    <%if valor4<>"" then%>
    <td width="17%" bgcolor="#FF9900"><font face="Verdana" size="1"><%=VALOR4%></font></td>
    <%ELSE%>
    <td width="17%"><font face="Verdana" size="1"><%=VALOR4%></font></td>
	<%END IF%>
    
    <%if valor5<>"" then%>
    <td width="17%" bgcolor="#FFFF00"><font face="Verdana" size="1"><%=valor5%></font></td>
    <%else%>
    <td width="17%"><font face="Verdana" size="1"><%=valor5%></font></td>
    <%end if%>
    
    <td width="17%" bgcolor="#AAFFDD"><font face="Verdana" size="1"><%=valor6%></font></td>
  </tr>
<%

mega_ant=RS("MEPR_CD_MEGA_PROCESSO")
proc_ant=RS("PROC_CD_PROCESSO")
sub_ant=RS("SUPR_CD_SUB_PROCESSO")
modulo_ant=RS("MODU_CD_MODULO")
atividade_ant=RS("ATCA_CD_ATIVIDADE_CARGA")

rs.movenext

ON ERROR RESUME NEXT

mega_atual=RS("MEPR_CD_MEGA_PROCESSO")
proc_atual=RS("PROC_CD_PROCESSO")
sub_atual=RS("SUPR_CD_SUB_PROCESSO")
modulo_atual=RS("MODU_CD_MODULO")
atividade_atual=RS("ATCA_CD_ATIVIDADE_CARGA")

LOOP
end if%>
</table>
<p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
</form>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
