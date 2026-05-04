 
<%
Response.Buffer = TRUE
Response.ContentType = "application/vnd.ms-excel"

server.scripttimeout=99999999

mega1=request("selMegaFuncao")

valor=request("selMegaProcesso")
proc=request("selProcesso")
subproc=request("selSubProcesso")

tatual=0

if proc<>0 then
	complemento=" AND PROC_CD_PROCESSO=" & proc
end if

if subproc<>0 then
	complemento=complemento+" AND SUPR_CD_SUB_PROCESSO=" & subproc
end if



ON ERROR RESUME NEXT

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & mega1 & " ORDER BY FUNE_CD_FUNCAO_NEGOCIO")
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

<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório de Associação de
Funções de Negócio</font></p>

<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#330099" size="2">Mega-Processo selecionado :
<%=valor%>-<%=texto%></font></b></p>

<%
cor=4

TEM=0

DO UNTIL RS.EOF=TRUE
	ssql1="SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & valor & " AND FUNE_CD_FUNCAO_NEGOCIO='" & RS("FUNE_CD_FUNCAO_NEGOCIO_PAI") & "'"+ complemento
	set ATUAL2=db.execute(ssql1)
	IF ATUAL2.EOF=FALSE THEN
		TEM=TEM+1
	END IF
RS.MOVENEXT	
LOOP

RS.MOVEFIRST

IF TEM<>0 THEN

set ATUAL=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & temp("MEPR_CD_MEGA_PROCESSO"))

if atual.eof=false then 

IF RS.EOF=FALSE THEN

tatual=1
%>
<p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp;

</font>

<p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp;

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" id="AutoNumber1" width="635">
  
  <tr>
  <td width="445" bgcolor="#330099" height="13" colspan="2" align="center">
  <p align="right" style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1" color="#FFFFFF">Fun&ccedil;&atilde;o R/3 --&gt;</font></b></td>
  <%
  do until rs.eof=true
  set functemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "'")
  tit_funcao=functemp("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
  %>
  <td width="98" height="35" rowspan="2" bgcolor="#FFFFFF" valign="middle">
    <p align="center" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><a href="#" onclick=javascript:window.open("exibe_funcao.asp?selFuncao=<%=rs("FUNE_CD_FUNCAO_NEGOCIO")%>","","width=550,height=340,status=0,toolbar=0")><%=tit_funcao%>  
    </font></a></p>
 </td>
  <%
  rs.movenext
  loop

  valor_sql="SELECT DISTINCT MEPR_CD_MEGA_PROCESSO, TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & valor + complemento
  set rs1=db.execute(valor_sql)
  %>
  </tr>
  
  <tr>
  <td width="281" bgcolor="#330099" height="20" align="center">
  <p style="margin-top: 0; margin-bottom: 0" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">Mega-Processo</font></b></p>
  </td>
  <td width="162" bgcolor="#330099" height="20" align="center">
  <p style="margin-top: 0; margin-bottom: 0" align="center">
  <b>
  <font size="1" face="Verdana" color="#FFFFFF">Transação</font></b></p>
  </td>
  </tr>
  <%
  MEGA_ANTERIOR=""
  
  DO UNTIL RS1.EOF=TRUE
  
  RS.MOVEFIRST
  
  EXISTE=0
    
  DO UNTIL RS.EOF=TRUE
  set rst=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RS1("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "' AND TRAN_CD_TRANSACAO='" & RS1("TRAN_CD_TRANSACAO") & "'")
  IF RST.EOF=FALSE THEN
  EXISTE=EXISTE+1
  END IF
  RS.MOVENEXT 
  LOOP
  IF EXISTE<>0 THEN 
  %>
  <tr>
    <td width="281" height="24" bgcolor="#FFFFFF">
      <%
      SET MEGA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs1("MEPR_CD_MEGA_PROCESSO"))
      VALOR_MEGA=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")
      IF TRIM(MEGA_ANTERIOR)=TRIM(VALOR_MEGA)THEN
      VALOR_ATUAL=""
      ELSE
      VALOR_ATUAL=VALOR_MEGA
      END IF
      MEGA_ANTERIOR=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")
      %>
      <p style="margin-top: 0; margin-bottom: 0" align="center"><font size="1" face="Verdana"><%=VALOR_ATUAL%></font></p>
    </td>
       <td width="162" height="24" bgcolor="#FFFFFF">
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
			COLOR="#C0C0C0"
	END SELECT	
	
   rs.movefirst 
	
	DO UNTIL RS.EOF=TRUE
	set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RS1("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "' AND TRAN_CD_TRANSACAO='" & RS1("TRAN_CD_TRANSACAO") & "'")
	IF RSTEMP.EOF=TRUE THEN
	%>
   <td width="98" height="24" bgcolor="<%=color%>">
      <p style="margin-top: 0; margin-bottom: 0" align="center"></td>
    <%
	ELSE
	%>
    <td width="85" height="24" bgcolor="<%=color%>">
      <p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="3" color="#330099" face="Informal011 BT">X</font></b></td>	
	<%
	 END IF
	 
	 RS.MOVENEXT
    LOOP
    
    END IF
    RS1.MOVENEXT
	 LOOP
	
	 END IF
	 END IF
	 END IF
    %>
  </tr>
 </table>
 <%IF TATUAL=0 THEN%>
<p><b><font face="Arial Unicode MS" color="#663300">Nenhum Registro Encontrado</font></b></p>
<%END IF%>
</body>

</html>






















