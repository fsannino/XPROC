 
<%
Response.Buffer = TRUE
Response.ContentType = "application/vnd.ms-excel"

server.scripttimeout=99999999

valor=request("selMegaProcesso")
tatual=0

ON ERROR RESUME NEXT

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & valor & " ORDER BY FUNE_CD_FUNCAO_NEGOCIO")

set rsorigem=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO ORDER BY MEPR_CD_MEGA_PROCESSO,PROC_CD_PROCESSO,SUPR_CD_SUB_PROCESSO")

set temp=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & valor)
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

<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório de Associação de
Funções de Negócio</font></p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#330099" size="2">Mega-Processo selecionado :
<%=valor%>-<%=texto%></font></b></p>

<%
cor=4
do until rsorigem.eof=true

TEM=0

DO UNTIL RS.EOF=TRUE
	ssql1="SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RSorigem("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & RS("FUNE_CD_FUNCAO_NEGOCIO") & "'"
	set ATUAL2=db.execute(ssql1)
	IF ATUAL2.EOF=FALSE THEN
		TEM=TEM+1
	END IF
RS.MOVENEXT	
LOOP

RS.MOVEFIRST

IF TEM<>0 THEN


set ATUAL=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RSorigem("MEPR_CD_MEGA_PROCESSO"))

if atual.eof=false then 

SET MEGA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rsorigem("MEPR_CD_MEGA_PROCESSO"))
VALOR=MEGA("MEPR_TX_dESC_MEGA_PROCESSO")

SET PROC=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rsorigem("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO=" & rsorigem("PROC_CD_PROCESSO"))
VALOR2=PROC("PROC_TX_DESC_PROCESSO")

IF RS.EOF=FALSE THEN
tatual=1
%>
<p align="left"><font face="Verdana" size="2"><b><font color="#0000CC">Mega-Processo</font>
</b><font size="1" face="Arial Unicode MS" color="#0000CC"> : <%=UCASE(VALOR)%></font>
</font>
<p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" size="2"><b><font color="#0000CC">Processo</font></b><font size="1" face="Arial Unicode MS" color="#0000CC">
: </font></font><font face="Verdana" size="2"><%=UCASE(VALOR2)%></font></p>

<p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" size="2"><b><font color="#0000CC">Sub-Processo</font></b><font size="1" face="Arial Unicode MS" color="#0000CC">
: </font></font></p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" width="642" height="76">
  
  <tr>
  <td width="476" bgcolor="#FFFFCC" height="17" colspan="2" align="center">
  <p align="right" style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1" color="#FF0000">Fun&ccedil;&atilde;o R/3 --&gt;</font></b></td>
  <%do until rs.eof=true
  set functemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "'")
  tit_funcao=functemp("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
  %>
  <td width="83" height="34" rowspan="2" bgcolor="#D6F8FC" valign="middle">
    <p align="center" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><a href="#" onclick=javascript:window.open("exibe_funcao.asp?selFuncao=<%=rs("FUNE_CD_FUNCAO_NEGOCIO")%>","","width=550,height=320,status=0,toolbar=0")><%=tit_funcao%>  
    </font></a></p>
 </td>
  <%
  rs.movenext
  loop
  rs.movefirst 
  set rs1=db.execute("SELECT DISTINCT ATCA_CD_ATIVIDADE_CARGA, TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RSorigem("MEPR_CD_MEGA_PROCESSO"))
  %>
  </tr>
  
  <tr>
  <td width="369" bgcolor="#FFFFCC" height="17" align="center">
  <p style="margin-top: 0; margin-bottom: 0" align="center">
  <b>
  <font face="Verdana" size="1">Atividade</font></b></p>
  </td>
  <td width="105" bgcolor="#FFFFCC" height="17" align="center">
  <p style="margin-top: 0; margin-bottom: 0" align="center">
  <b>
  <font size="1" face="Verdana" color="#0000CC">Transação</font></b></p>
  </td>
  </tr>
  <%
  DO UNTIL RS1.EOF=TRUE
  SET AT=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & RS1("ATCA_CD_ATIVIDADE_CARGA"))
  ATIVIDADE=AT("ATCA_TX_DESC_ATIVIDADE")
  %>
  <tr>
    <td width="369" height="21" bgcolor="#EAD997">
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1">
      <%=ATIVIDADE%></font></td>
       <td width="105" height="21" bgcolor="#DACFF1">
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" size="1"><%=RS1("TRAN_CD_TRANSACAO")%></font></td>
    <%
    IF COR=1 THEN
		COR=4
	ELSE
		COR=1
	END IF
	
	SELECT CASE COR
		CASE 1
			COLOR="#FAF4D8"
		CASE 4
			color="#C0C0C0"
	END SELECT	
	
   DO UNTIL RS.EOF=TRUE
	set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RSorigem("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "' AND TRAN_CD_TRANSACAO='" & RS1("TRAN_CD_TRANSACAO") & "'")
   IF RSTEMP.EOF=TRUE THEN
	%>
   <td width="83" height="21" bgcolor="<%=color%>">
      <p style="margin-top: 0; margin-bottom: 0" align="center"></td>
    <%
	ELSE
	%>
    <td width="78" height="21" bgcolor="<%=color%>">
      <p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="3" color="#330099" face="Informal011 BT">X</font></b></td>	
	<%
	END IF
    RS.MOVENEXT
    LOOP
    %>
  </tr>
  <%
  RS1.MOVENEXT
  RS.MOVEFIRST
  LOOP
  END IF
  %>
 </table>
 <%
 end if
 END IF
 RSORIGEM.MOVENEXT
 loop
 if tatual=0 then
 %>
<p><b><font face="Arial Unicode MS" color="#663300">Nenhum Registro Encontrado</font></b></p>
<%end if%>
</body>

</html>