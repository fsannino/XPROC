<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

MODULO=REQUEST("selModulo")
ATIVIDADE=REQUEST("selAtividade")
TRANSACAO=REQUEST("selTransacao")

IF MODULO<>0 THEN
COMPL1="MODU_CD_MODULO=" & MODULO
END IF

IF ATIVIDADE<>0 THEN
COMPL2="ATCA_CD_ATIVIDADE_CARGA=" & ATIVIDADE
END IF

IF TRANSACAO<>"0" THEN
COMPL3="TRAN_CD_TRANSACAO='" & TRANSACAO & "'"
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

IF COMPLE<>"" THEN
CONECTA="WHERE "
END IF

ORDENA=" ORDER BY MODU_CD_MODULO, ATCA_CD_ATIVIDADE_CARGA, TRAN_CD_TRANSACAO"
SSQL="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA " & CONECTA & COMPLE & ORDENA

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
            <td width="195"><img src="../imagens/print.gif" width="90" height="35" border="0" onclick="javascript:print()"></td>
          <td width="27"></td>
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
  Agrupamento de Processo x Atividade x Transação</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
<%if rs.eof=true then%>  
<p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" color="#800000" size="2"><b>Nenhum
Registro Encontrado</b></font></p>
<%end if%>

<%if rs.eof=false then%>
<table border="0" width="100%" cellspacing="3">
  <tr>
    <td width="33%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Agrupamento
      das Atividades</b></font></td>
    <td width="33%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Atividade</b></font></td>
    <td width="34%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Transação</b></font></td>
  </tr>
  <%
  IF RS.EOF<>TRUE THEN
  VALOR_MODULO=RS("MODU_CD_MODULO")
  VALOR_ATIVIDADE=RS("ATCA_CD_ATIVIDADE_CARGA")
  VALOR_TRANSACAO=RS("TRAN_CD_TRANSACAO")
  END IF
  
  VALOR1=""
  VALOR2=""
  VALOR3=""

  DO UNTIL RS.EOF=TRUE
  
  'IF MODULO_ANTERIOR<>VALOR_MODULO THEN
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & VALOR_MODULO)
  VALOR1=RS1("MODU_TX_DESC_MODULO")
  'END IF
  
  IF MODULO_ANTERIOR<>VALOR_MODULO or ATIVIDADE_ANTERIOR <> VALOR_ATIVIDADE THEN
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & VALOR_ATIVIDADE)
  VALOR2=RS1("ATCA_TX_DESC_ATIVIDADE")
  END IF
  
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & VALOR_TRANSACAO & "'")
  VALOR3=RS("TRAN_CD_TRANSACAO") & "-"& RS1("TRAN_TX_DESC_TRANSACAO")
  
  %>
  <tr>
    <%if valor2<>"" then%>
    <td width="33%" bgcolor="#FFCC00"><font face="Verdana" size="1"><%=VALOR1%></font></td>
    <%else%>
    <td width="33%"><font face="Verdana" size="1"></font></td>
    <%end if%>
	
    <%if valor2<>"" then%>
    <td width="33%" bgcolor="#FFFF00"><font face="Verdana" size="1"><%=VALOR2%></font></td>
	<%else%>
	<td width="33%"><font face="Verdana" size="1"><%=VALOR2%></font></td>
    <%end if%>
    
    <%if valor3<>"" then%>
    <td width="34%" bgcolor="#FFCAB0"><font face="Verdana" size="1"><%=VALOR3%></font></td>
    <%else%>
    <td width="34%"><font face="Verdana" size="1"><%=VALOR3%></font></td>

    <%end if%>
  </tr>
  <%
  
  MODULO_ANTERIOR=RS("MODU_CD_MODULO")
  ATIVIDADE_ANTERIOR=RS("ATCA_CD_ATIVIDADE_CARGA")
  
  RS.MOVENEXT
  
  IF RS.EOF<>TRUE THEN
  VALOR_MODULO=RS("MODU_CD_MODULO")
  VALOR_ATIVIDADE=RS("ATCA_CD_ATIVIDADE_CARGA")
  VALOR_TRANSACAO=RS("TRAN_CD_TRANSACAO")
  END IF
  
  VALOR1=""
  VALOR2=""
  VALOR3=""
  
  LOOP
  %>
</table>
<%end if%>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
</form>

</body>

</html>
=======
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

MODULO=REQUEST("selModulo")
ATIVIDADE=REQUEST("selAtividade")
TRANSACAO=REQUEST("selTransacao")

IF MODULO<>0 THEN
COMPL1="MODU_CD_MODULO=" & MODULO
END IF

IF ATIVIDADE<>0 THEN
COMPL2="ATCA_CD_ATIVIDADE_CARGA=" & ATIVIDADE
END IF

IF TRANSACAO<>"0" THEN
COMPL3="TRAN_CD_TRANSACAO='" & TRANSACAO & "'"
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

IF COMPLE<>"" THEN
CONECTA="WHERE "
END IF

ORDENA=" ORDER BY MODU_CD_MODULO, ATCA_CD_ATIVIDADE_CARGA, TRAN_CD_TRANSACAO"
SSQL="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA " & CONECTA & COMPLE & ORDENA

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
            <td width="195"><img src="../imagens/print.gif" width="90" height="35" border="0" onclick="javascript:print()"></td>
          <td width="27"></td>
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
  Agrupamento de Processo x Atividade x Transação</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
<%if rs.eof=true then%>  
<p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" color="#800000" size="2"><b>Nenhum
Registro Encontrado</b></font></p>
<%end if%>

<%if rs.eof=false then%>
<table border="0" width="100%" cellspacing="3">
  <tr>
    <td width="33%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Agrupamento
      das Atividades</b></font></td>
    <td width="33%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Atividade</b></font></td>
    <td width="34%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Transação</b></font></td>
  </tr>
  <%
  IF RS.EOF<>TRUE THEN
  VALOR_MODULO=RS("MODU_CD_MODULO")
  VALOR_ATIVIDADE=RS("ATCA_CD_ATIVIDADE_CARGA")
  VALOR_TRANSACAO=RS("TRAN_CD_TRANSACAO")
  END IF
  
  VALOR1=""
  VALOR2=""
  VALOR3=""

  DO UNTIL RS.EOF=TRUE
  
  'IF MODULO_ANTERIOR<>VALOR_MODULO THEN
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & VALOR_MODULO)
  VALOR1=RS1("MODU_TX_DESC_MODULO")
  'END IF
  
  IF MODULO_ANTERIOR<>VALOR_MODULO or ATIVIDADE_ANTERIOR <> VALOR_ATIVIDADE THEN
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & VALOR_ATIVIDADE)
  VALOR2=RS1("ATCA_TX_DESC_ATIVIDADE")
  END IF
  
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & VALOR_TRANSACAO & "'")
  VALOR3=RS("TRAN_CD_TRANSACAO") & "-"& RS1("TRAN_TX_DESC_TRANSACAO")
  
  %>
  <tr>
    <%if valor2<>"" then%>
    <td width="33%" bgcolor="#FFCC00"><font face="Verdana" size="1"><%=VALOR1%></font></td>
    <%else%>
    <td width="33%"><font face="Verdana" size="1"></font></td>
    <%end if%>
	
    <%if valor2<>"" then%>
    <td width="33%" bgcolor="#FFFF00"><font face="Verdana" size="1"><%=VALOR2%></font></td>
	<%else%>
	<td width="33%"><font face="Verdana" size="1"><%=VALOR2%></font></td>
    <%end if%>
    
    <%if valor3<>"" then%>
    <td width="34%" bgcolor="#FFCAB0"><font face="Verdana" size="1"><%=VALOR3%></font></td>
    <%else%>
    <td width="34%"><font face="Verdana" size="1"><%=VALOR3%></font></td>

    <%end if%>
  </tr>
  <%
  
  MODULO_ANTERIOR=RS("MODU_CD_MODULO")
  ATIVIDADE_ANTERIOR=RS("ATCA_CD_ATIVIDADE_CARGA")
  
  RS.MOVENEXT
  
  IF RS.EOF<>TRUE THEN
  VALOR_MODULO=RS("MODU_CD_MODULO")
  VALOR_ATIVIDADE=RS("ATCA_CD_ATIVIDADE_CARGA")
  VALOR_TRANSACAO=RS("TRAN_CD_TRANSACAO")
  END IF
  
  VALOR1=""
  VALOR2=""
  VALOR3=""
  
  LOOP
  %>
</table>
<%end if%>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
</form>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
