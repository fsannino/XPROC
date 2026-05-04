<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->

<%
SERVER.SCRIPTTIMEOUT=99999999

num_mega=request("selMegaProcesso")
num_processo=request("selProcesso")
num_sub=request("selSubProcesso")

nume_mega=num_mega
nume_proc=0
nume_sub=0
nume_ativ=0

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if num_processo<>0 and num_sub<>0 then
		sql_compl="WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO="& num_mega &" AND " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO="& num_processo &" AND " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO=" & num_sub
	else
	if num_processo<>0 then
		sql_compl="WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO="& num_mega &" AND " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO="& num_processo 
	else
		sql_compl="WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO="& num_mega
	end if
end if

sql="SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO FROM (((" & Session("PREFIXO") & "MEGA_PROCESSO INNER JOIN " & Session("PREFIXO") & "PROCESSO ON " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO)  INNER JOIN " & Session("PREFIXO") & "SUB_PROCESSO ON (" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO)) INNER JOIN " & Session("PREFIXO") & "RELACAO_FINAL ON (" & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO) AND (" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO))  INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA "
ssql=sql+sql_compl+"GROUP BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA"

set rs=db.execute(ssql)
%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0">
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
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
          <td width="50"><a href="javascript:print()">
    <img border="0" src="../imagens/print.gif" align="left">
</a></td>
          <td width="26">&nbsp;</td>
          <td width="195">
			<%
				mega=0
				processo=0			
				subr=0
				mega=num_mega
				processo=num_processo
				if len(num_sub)<>0 then
				subr=num_sub
				end if
			%>          
          <p align="center"><a href="resultado_consulta_excel.asp?selMegaProcesso=<%=mega%>&amp;selProcesso=<%=processo%>&amp;selSubProcesso=<%=subr%>" target="blank"><img border="0" src="../imagens/exp_excel.gif"></a></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;

</p>
<p style="margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana" color="#330099">Resultado
da Consulta</font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;

</p>
<p style="margin-top: 0; margin-bottom: 0">
<%if rs.eof=true then%>
<b><font size="2" color="#800000" face="Verdana"> Nenhum Registro Encontrado </font></b>
<%
end if

on error resume next

mega_atual=rs("MEPR_CD_MEGA_PROCESSO")
proc_atual=rs("PROC_CD_PROCESSO")
sub_atual=rs("SUPR_CD_SUB_PROCESSO")

do until rs.eof=true

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual)
set rs_processo=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual & " AND PROC_CD_PROCESSO="& proc_atual)
set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual & " AND PROC_CD_PROCESSO="& proc_atual & " AND SUPR_CD_SUB_PROCESSO="& sub_atual )

%>
</p>
<table border="0" width="639">
  <%if mega_atual<>mega_ant then
  nume_mega=mega_atual
  nume_proc=0
  nume_sub=0
  nume_ativ=0
  %>
  <tr bgcolor="#FF9933"> 
    <td width="161"><font face="Verdana" size="2">MEGA-PROCESSO</font></td>
    <td width="36" align="left"><b><font face="Verdana" size=2><%=nume_mega%></font></b></td>
    <td width="422"><b><font face="Verdana" size=2><%=rs_mega("MEPR_TX_DESC_MEGA_PROCESSO")%></font></b></td>
  </tr>
  <%end if%>
  <%if proc_atual<>proc_ant then
  nume_proc=nume_proc+1
  nume_sub=0
  nume_ativ=0
  %>
  <tr bgcolor="#FFCC66"> 
    <td width="161"><font face="Verdana" size="2">PROCESSO</font></td>
    <td width="36" align="left"><b><font face="Verdana" size=2><%=nume_mega & "." & nume_proc%></font></b></td>
    <td width="422"><b><font face="Verdana" size=2><%=rs_processo("PROC_TX_DESC_PROCESSO")%></font></b></td>
  </tr>
  <%end if
  nume_sub=nume_sub+1
  %>
  <tr bgcolor="#FFFFCC"> 
    <td width="161"><font face="Verdana" size="2">SUB-PROCESSO</font></td>
    <td width="36" align="left"><b><font face="Verdana" size=2><%=nume_mega & "." & nume_proc & "." & nume_sub%></font></b></td>
    <td width="422" bgcolor="#FFFFCC"> 
      <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size=2><%=rs_sub("SUPR_TX_DESC_SUB_PROCESSO")%></font></b> 
    </td>
  </tr>
</table>

<div align="left">

  <table border="0" width="638" cellpadding="0" cellspacing="0" height="55">
    <%

ssql="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual &" AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual & " ORDER BY ATCA_CD_ATIVIDADE_CARGA, RELA_NR_SEQUENCIA"

set rs_ativ=db.execute(ssql)

on error resume next
ativ_atual=rs_ativ("ATCA_CD_ATIVIDADE_CARGA")
trans_atual=rs_ativ("TRAN_CD_TRANSACAO")
seq_atual=rs_ativ("ATCA_NR_SEQUENCIA")

%>
    <%if rs_ativ.eof=false then%>
    <tr> 
      <td width="152" bgcolor="#CCCCCC" align="left" height="20"><font face="Verdana" size="2">ATIVIDADE</font></td>
      <td width="320" bgcolor="#CCCCCC" align="left" height="20"><b><font face="Verdana" size="2">DESCRIÇÃO 
        TRANSA&Ccedil;&Atilde;O </font></b></td>
      <td width="148" bgcolor="#CCCCCC" align="left" height="20"> 
        <p align="left"><b><font face="Verdana" size="2"> TRANSA&Ccedil;&Atilde;O</font></b></p>
      </td>
    </tr>
    <%else
nenhum=1
%>
      <td width="152" height="32"> <font size="2" color="#800000" face="Verdana"> 
        Nenhum registro encontrado </font>  
        <%
end if

ativ_anterior=0

nume_ativ=0

do until rs_ativ.eof=true

%>
    <tr> 
      <%if ativ_anterior = ativ_atual  then%>
      <td width="152" height="21"><font face="Arial" size="2">&nbsp;</font></td>
      <%else
	nume_ativ=nume_ativ+1
    set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA="& ativ_atual)
    PREFIXO=nume_mega & "." & nume_proc & "." & nume_sub & "." & nume_ativ
    ATIVIDADE=rs1("ATCA_TX_DESC_ATIVIDADE")
	%>
      <td width="320" height="21"><font face="Arial" size="1"><%=PREFIXO%> - <%=ATIVIDADE%></font></td>
      <%end if%>
      <%
    set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='"& trim(trans_atual) & "'")
    TRANSACAO = rs2("TRAN_TX_DESC_TRANSACAO")
    %>
      <td width="148" height="21"><font face="Arial" size="1" align="left"> 
        <p align="left"><%=TRANSACAO%>
        </font></td>
      <td width="18" height="21"><font face="Arial" size="1" align="right"> 
        <p align="left"><%=trans_atual%>
        </font></td>
    </tr>
    <%
trans_anterior=rs_ativ("TRAN_CD_TRANSACAO")
ativ_anterior=rs_ativ("ATCA_CD_ATIVIDADE_CARGA")
seq_anterior=rs_ativ("ATCA_NR_SEQUENCIA")
rs_ativ.movenext
on error resume next
trans_atual=rs_ativ("TRAN_CD_TRANSACAO")
ativ_atual=rs_ativ("ATCA_CD_ATIVIDADE_CARGA")
seq_atual=rs_ativ("ATCA_NR_SEQUENCIA")
loop
%>
  </table>

</div>

<%
mega_ant=rs("MEPR_CD_MEGA_PROCESSO")
proc_ant=rs("PROC_CD_PROCESSO")
sub_ant=rs("SUPR_CD_SUB_PROCESSO")

rs.movenext

on error resume next

mega_atual=rs("MEPR_CD_MEGA_PROCESSO")
proc_atual=rs("PROC_CD_PROCESSO")
sub_atual=rs("SUPR_CD_SUB_PROCESSO")

loop

%>

<p align="center">
<br>

</body>

</html>


=======
 
<!--#include file="../asp/protege/protege.asp" -->

<%
SERVER.SCRIPTTIMEOUT=99999999

num_mega=request("selMegaProcesso")
num_processo=request("selProcesso")
num_sub=request("selSubProcesso")

nume_mega=num_mega
nume_proc=0
nume_sub=0
nume_ativ=0

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if num_processo<>0 and num_sub<>0 then
		sql_compl="WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO="& num_mega &" AND " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO="& num_processo &" AND " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO=" & num_sub
	else
	if num_processo<>0 then
		sql_compl="WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO="& num_mega &" AND " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO="& num_processo 
	else
		sql_compl="WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO="& num_mega
	end if
end if

sql="SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO FROM (((" & Session("PREFIXO") & "MEGA_PROCESSO INNER JOIN " & Session("PREFIXO") & "PROCESSO ON " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO)  INNER JOIN " & Session("PREFIXO") & "SUB_PROCESSO ON (" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO)) INNER JOIN " & Session("PREFIXO") & "RELACAO_FINAL ON (" & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO) AND (" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO))  INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA "
ssql=sql+sql_compl+"GROUP BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA"

set rs=db.execute(ssql)
%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0">
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
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
          <td width="50"><a href="javascript:print()">
    <img border="0" src="../imagens/print.gif" align="left">
</a></td>
          <td width="26">&nbsp;</td>
          <td width="195">
			<%
				mega=0
				processo=0			
				subr=0
				mega=num_mega
				processo=num_processo
				if len(num_sub)<>0 then
				subr=num_sub
				end if
			%>          
          <p align="center"><a href="resultado_consulta_excel.asp?selMegaProcesso=<%=mega%>&amp;selProcesso=<%=processo%>&amp;selSubProcesso=<%=subr%>" target="blank"><img border="0" src="../imagens/exp_excel.gif"></a></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;

</p>
<p style="margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana" color="#330099">Resultado
da Consulta</font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;

</p>
<p style="margin-top: 0; margin-bottom: 0">
<%if rs.eof=true then%>
<b><font size="2" color="#800000" face="Verdana"> Nenhum Registro Encontrado </font></b>
<%
end if

on error resume next

mega_atual=rs("MEPR_CD_MEGA_PROCESSO")
proc_atual=rs("PROC_CD_PROCESSO")
sub_atual=rs("SUPR_CD_SUB_PROCESSO")

do until rs.eof=true

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual)
set rs_processo=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual & " AND PROC_CD_PROCESSO="& proc_atual)
set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual & " AND PROC_CD_PROCESSO="& proc_atual & " AND SUPR_CD_SUB_PROCESSO="& sub_atual )

%>
</p>
<table border="0" width="639">
  <%if mega_atual<>mega_ant then
  nume_mega=mega_atual
  nume_proc=0
  nume_sub=0
  nume_ativ=0
  %>
  <tr bgcolor="#FF9933"> 
    <td width="161"><font face="Verdana" size="2">MEGA-PROCESSO</font></td>
    <td width="36" align="left"><b><font face="Verdana" size=2><%=nume_mega%></font></b></td>
    <td width="422"><b><font face="Verdana" size=2><%=rs_mega("MEPR_TX_DESC_MEGA_PROCESSO")%></font></b></td>
  </tr>
  <%end if%>
  <%if proc_atual<>proc_ant then
  nume_proc=nume_proc+1
  nume_sub=0
  nume_ativ=0
  %>
  <tr bgcolor="#FFCC66"> 
    <td width="161"><font face="Verdana" size="2">PROCESSO</font></td>
    <td width="36" align="left"><b><font face="Verdana" size=2><%=nume_mega & "." & nume_proc%></font></b></td>
    <td width="422"><b><font face="Verdana" size=2><%=rs_processo("PROC_TX_DESC_PROCESSO")%></font></b></td>
  </tr>
  <%end if
  nume_sub=nume_sub+1
  %>
  <tr bgcolor="#FFFFCC"> 
    <td width="161"><font face="Verdana" size="2">SUB-PROCESSO</font></td>
    <td width="36" align="left"><b><font face="Verdana" size=2><%=nume_mega & "." & nume_proc & "." & nume_sub%></font></b></td>
    <td width="422" bgcolor="#FFFFCC"> 
      <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size=2><%=rs_sub("SUPR_TX_DESC_SUB_PROCESSO")%></font></b> 
    </td>
  </tr>
</table>

<div align="left">

  <table border="0" width="638" cellpadding="0" cellspacing="0" height="55">
    <%

ssql="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual &" AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual & " ORDER BY ATCA_CD_ATIVIDADE_CARGA, RELA_NR_SEQUENCIA"

set rs_ativ=db.execute(ssql)

on error resume next
ativ_atual=rs_ativ("ATCA_CD_ATIVIDADE_CARGA")
trans_atual=rs_ativ("TRAN_CD_TRANSACAO")
seq_atual=rs_ativ("ATCA_NR_SEQUENCIA")

%>
    <%if rs_ativ.eof=false then%>
    <tr> 
      <td width="152" bgcolor="#CCCCCC" align="left" height="20"><font face="Verdana" size="2">ATIVIDADE</font></td>
      <td width="320" bgcolor="#CCCCCC" align="left" height="20"><b><font face="Verdana" size="2">DESCRIÇÃO 
        TRANSA&Ccedil;&Atilde;O </font></b></td>
      <td width="148" bgcolor="#CCCCCC" align="left" height="20"> 
        <p align="left"><b><font face="Verdana" size="2"> TRANSA&Ccedil;&Atilde;O</font></b></p>
      </td>
    </tr>
    <%else
nenhum=1
%>
      <td width="152" height="32"> <font size="2" color="#800000" face="Verdana"> 
        Nenhum registro encontrado </font>  
        <%
end if

ativ_anterior=0

nume_ativ=0

do until rs_ativ.eof=true

%>
    <tr> 
      <%if ativ_anterior = ativ_atual  then%>
      <td width="152" height="21"><font face="Arial" size="2">&nbsp;</font></td>
      <%else
	nume_ativ=nume_ativ+1
    set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA="& ativ_atual)
    PREFIXO=nume_mega & "." & nume_proc & "." & nume_sub & "." & nume_ativ
    ATIVIDADE=rs1("ATCA_TX_DESC_ATIVIDADE")
	%>
      <td width="320" height="21"><font face="Arial" size="1"><%=PREFIXO%> - <%=ATIVIDADE%></font></td>
      <%end if%>
      <%
    set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='"& trim(trans_atual) & "'")
    TRANSACAO = rs2("TRAN_TX_DESC_TRANSACAO")
    %>
      <td width="148" height="21"><font face="Arial" size="1" align="left"> 
        <p align="left"><%=TRANSACAO%>
        </font></td>
      <td width="18" height="21"><font face="Arial" size="1" align="right"> 
        <p align="left"><%=trans_atual%>
        </font></td>
    </tr>
    <%
trans_anterior=rs_ativ("TRAN_CD_TRANSACAO")
ativ_anterior=rs_ativ("ATCA_CD_ATIVIDADE_CARGA")
seq_anterior=rs_ativ("ATCA_NR_SEQUENCIA")
rs_ativ.movenext
on error resume next
trans_atual=rs_ativ("TRAN_CD_TRANSACAO")
ativ_atual=rs_ativ("ATCA_CD_ATIVIDADE_CARGA")
seq_atual=rs_ativ("ATCA_NR_SEQUENCIA")
loop
%>
  </table>

</div>

<%
mega_ant=rs("MEPR_CD_MEGA_PROCESSO")
proc_ant=rs("PROC_CD_PROCESSO")
sub_ant=rs("SUPR_CD_SUB_PROCESSO")

rs.movenext

on error resume next

mega_atual=rs("MEPR_CD_MEGA_PROCESSO")
proc_atual=rs("PROC_CD_PROCESSO")
sub_atual=rs("SUPR_CD_SUB_PROCESSO")

loop

%>

<p align="center">
<br>

</body>

</html>


>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
