<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

num_mega=request("selMegaProcesso")
str_emp=request("txtEmpSelecionada")

str_valor = str_emp

if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if

tamanho = Len(str_valor)

If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If

tamanho = Len(str_valor)
contador = 1

Do Until contador = tamanho + 1

    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    
    If str_temp = "," Then
       
       str_atual = Right(str_atual, quantos)
       str_atual = Left(str_atual, quantos - 1)
        
		'Aqui entra o que vc quer fazer com o caracter em questão!
		
		if len(compl_emp)=0 then
		compl_emp = compl_emp & "(" & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.EMPR_CD_NR_EMPRESA=" & str_atual & ")"
		else
		compl_emp = compl_emp & " OR (" & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.EMPR_CD_NR_EMPRESA=" & str_atual & ")"
		end if

		set rs_empresa=db.execute("SELECT * FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE WHERE EMPR_CD_NR_EMPRESA=" & str_atual)
		valor_emp=rs_empresa("EMPR_TX_NOME_EMPRESA")
		
		if len(str_empresas)=0 then
		    str_empresas=valor_emp
		else
			str_empresas = str_empresas & " / " &  valor_emp
		end if
		
		quantos=0
		
		conta_emp=conta_emp+1
		
    End If
       
    contador = contador + 1

Loop

nume_mega=num_mega
nume_proc=0
nume_sub=0
nume_ativ=0

if num_mega<>0 then
		sql_compl="WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO="& num_mega
end if

sql="SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO FROM (((" & Session("PREFIXO") & "MEGA_PROCESSO INNER JOIN " & Session("PREFIXO") & "PROCESSO ON " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO)  INNER JOIN " & Session("PREFIXO") & "SUB_PROCESSO ON (" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO)) INNER JOIN " & Session("PREFIXO") & "RELACAO_FINAL ON (" & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO) AND (" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO))  INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA "
ssql=sql+sql_compl+"GROUP BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA"

set rs=db.execute(ssql)

if nume_mega=0 then
	nume_mega=rs("MEPR_CD_MEGA_PROCESSO")
end if

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
          <td width="195"></td>
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
<table border="0" width="82%">
  <tr> 
    <td width="61%" height="36"> 
      <p style="margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana" color="#330099">Resultado 
        da Consulta</font></p>
      <%IF LEN(STR_EMPRESAS)>0 THEN%>
      <p style="margin-top: 0; margin-bottom: 0"><b><font color="#330099" face="Verdana" size="2">Empresas 
        / Unidades Selecionadas&nbsp;:</font></b></p>
      <p style="margin-top: 0; margin-bottom: 0"><b><font color="#330099" face="Verdana" size="1"> 
        </font></b></p>
      <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
      
  </tr>
  <tr>
    <td width="61%"> 
      <div align="center"><b><font color="#330099" face="Verdana" size="3"><%=str_empresas%></font></b> 
        <%END IF%>
      </div>
    </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">
<%if rs.eof=true then%>
</p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;

</p>
<p style="margin-top: 0; margin-bottom: 0">
&nbsp;
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
  
  nume_mega=num_mega
  
  if nume_mega=0 then
	nume_mega=rs("MEPR_CD_MEGA_PROCESSO")
  end if

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
    <td width="422"><b><font face="Verdana" size=2><%=ucase(rs_processo("PROC_TX_DESC_PROCESSO"))%></font></b></td>
  </tr>
  <%end if
  nume_sub=nume_sub+1
  %>
  <tr bgcolor="#FFFFCC"> 
    <td width="161"><font face="Verdana" size="2">SUB-PROCESSO</font></td>
    <td width="36" align="left"><b><font face="Verdana" size=2><%=nume_mega & "." & nume_proc & "." & nume_sub%></font></b></td>
    <td width="422"> 
      <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size=2><%=rs_sub("SUPR_TX_DESC_SUB_PROCESSO")%></font></b>
    </td>
  </tr>
</table>

<div align="left">

  <table border="0" width="639" cellpadding="0" cellspacing="0" height="43">
    <%

ssql01="SELECT " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO, " & Session("PREFIXO") & "RELACAO_FINAL.RELA_NR_SEQUENCIA, " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.EMPR_CD_NR_EMPRESA FROM " & Session("PREFIXO") & "RELACAO_FINAL INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.ATCA_CD_ATIVIDADE_CARGA"
sql_complemento01=" WHERE " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO="& mega_atual &" AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO=" & proc_atual & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO=" & sub_atual
ssql02=" ORDER BY " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO, " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.EMPR_CD_NR_EMPRESA"

if len(compl_emp)>0 then
ssql=ssql01+sql_complemento01+ " AND(" + compl_emp + ")" +ssql02
else
ssql="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual &" AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual & " ORDER BY ATCA_CD_ATIVIDADE_CARGA, RELA_NR_SEQUENCIA"
end if

'response.write ssql

set rs_ativ=db.execute(ssql)

on error resume next
ativ_atual=rs_ativ("ATCA_CD_ATIVIDADE_CARGA")
trans_atual=rs_ativ("TRAN_CD_TRANSACAO")
seq_atual=rs_ativ("ATCA_NR_SEQUENCIA")
%>
    <%if rs_ativ.eof=false then%>
    <tr> 
      <td width="157" bgcolor="#CCCCCC" align="left" height="20"><font face="Verdana" size="2">ATIVIDADE</font></td>
      <td width="215" bgcolor="#CCCCCC" align="left" height="20"><b><font face="Verdana" size="2">DESCRIÇÃO 
        TRANSA&Ccedil;&Atilde;O </font></b></td>
      <td width="232" bgcolor="#CCCCCC" align="left" height="20"> 
        <p align="left"><b><font face="Verdana" size="2"> TRANSA&Ccedil;&Atilde;O</font></b></p>
      </td>
    </tr>
    <%else
nenhum=1
%>
      <td width="157" height="32"> <font size="2" color="#800000" face="Verdana"> 
        Nenhum registro encontrado </font> </b> 
        <%
end if

ativ_anterior=0

nume_ativ=0

do until rs_ativ.eof=true
nova_ativ=0
%>
    <tr> 
      <%if ativ_anterior <> ativ_atual  then
   nova_ativ=1
	nume_ativ=nume_ativ+1
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA="& ativ_atual)
	PREFIXO=nume_mega & "." & nume_proc & "." & nume_sub & "." & nume_ativ
	ATIVIDADE=rs1("ATCA_TX_DESC_ATIVIDADE")
	%>
      <td width="157" height="21"><font face="Arial" size="1"><%=PREFIXO%> - <%=ATIVIDADE%></font></td>
      <%
   existe=1
   end if%>
      <%IF (TRANS_ANTERIOR<>TRANS_ATUAL) THEN
    set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='"& trim(trans_atual) & "'")
    TRANSACAO = rs2("TRAN_TX_DESC_TRANSACAO")
    if nova_ativ=0 then%>
      <td width="215" height="21"><font face="Arial" size="2">&nbsp;</font></td>
      <%end if%>
      <%IF nova_ativ=1 or TRANS_ANTERIOR<>TRANS_ATUAL THEN%>
      <td width="232" height="21"><font face="Arial" size="1" align="left"> 
        <p align="left"><%=TRANSACAO%>
        </font></td>
      <td width="31" height="21"><font face="Arial" size="1" align="right"> 
        <p align="left"><%=trans_atual%>
        </font></td>
      <%END IF
      END IF
      %>
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

existe=0

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
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

num_mega=request("selMegaProcesso")
str_emp=request("txtEmpSelecionada")

str_valor = str_emp

if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if

tamanho = Len(str_valor)

If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If

tamanho = Len(str_valor)
contador = 1

Do Until contador = tamanho + 1

    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    
    If str_temp = "," Then
       
       str_atual = Right(str_atual, quantos)
       str_atual = Left(str_atual, quantos - 1)
        
		'Aqui entra o que vc quer fazer com o caracter em questão!
		
		if len(compl_emp)=0 then
		compl_emp = compl_emp & "(" & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.EMPR_CD_NR_EMPRESA=" & str_atual & ")"
		else
		compl_emp = compl_emp & " OR (" & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.EMPR_CD_NR_EMPRESA=" & str_atual & ")"
		end if

		set rs_empresa=db.execute("SELECT * FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE WHERE EMPR_CD_NR_EMPRESA=" & str_atual)
		valor_emp=rs_empresa("EMPR_TX_NOME_EMPRESA")
		
		if len(str_empresas)=0 then
		    str_empresas=valor_emp
		else
			str_empresas = str_empresas & " / " &  valor_emp
		end if
		
		quantos=0
		
		conta_emp=conta_emp+1
		
    End If
       
    contador = contador + 1

Loop

nume_mega=num_mega
nume_proc=0
nume_sub=0
nume_ativ=0

if num_mega<>0 then
		sql_compl="WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO="& num_mega
end if

sql="SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO FROM (((" & Session("PREFIXO") & "MEGA_PROCESSO INNER JOIN " & Session("PREFIXO") & "PROCESSO ON " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO)  INNER JOIN " & Session("PREFIXO") & "SUB_PROCESSO ON (" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO)) INNER JOIN " & Session("PREFIXO") & "RELACAO_FINAL ON (" & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO) AND (" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO))  INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA "
ssql=sql+sql_compl+"GROUP BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA"

set rs=db.execute(ssql)

if nume_mega=0 then
	nume_mega=rs("MEPR_CD_MEGA_PROCESSO")
end if

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
          <td width="195"></td>
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
<table border="0" width="82%">
  <tr> 
    <td width="61%" height="36"> 
      <p style="margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana" color="#330099">Resultado 
        da Consulta</font></p>
      <%IF LEN(STR_EMPRESAS)>0 THEN%>
      <p style="margin-top: 0; margin-bottom: 0"><b><font color="#330099" face="Verdana" size="2">Empresas 
        / Unidades Selecionadas&nbsp;:</font></b></p>
      <p style="margin-top: 0; margin-bottom: 0"><b><font color="#330099" face="Verdana" size="1"> 
        </font></b></p>
      <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
      
  </tr>
  <tr>
    <td width="61%"> 
      <div align="center"><b><font color="#330099" face="Verdana" size="3"><%=str_empresas%></font></b> 
        <%END IF%>
      </div>
    </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">
<%if rs.eof=true then%>
</p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;

</p>
<p style="margin-top: 0; margin-bottom: 0">
&nbsp;
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
  
  nume_mega=num_mega
  
  if nume_mega=0 then
	nume_mega=rs("MEPR_CD_MEGA_PROCESSO")
  end if

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
    <td width="422"><b><font face="Verdana" size=2><%=ucase(rs_processo("PROC_TX_DESC_PROCESSO"))%></font></b></td>
  </tr>
  <%end if
  nume_sub=nume_sub+1
  %>
  <tr bgcolor="#FFFFCC"> 
    <td width="161"><font face="Verdana" size="2">SUB-PROCESSO</font></td>
    <td width="36" align="left"><b><font face="Verdana" size=2><%=nume_mega & "." & nume_proc & "." & nume_sub%></font></b></td>
    <td width="422"> 
      <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size=2><%=rs_sub("SUPR_TX_DESC_SUB_PROCESSO")%></font></b>
    </td>
  </tr>
</table>

<div align="left">

  <table border="0" width="639" cellpadding="0" cellspacing="0" height="43">
    <%

ssql01="SELECT " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO, " & Session("PREFIXO") & "RELACAO_FINAL.RELA_NR_SEQUENCIA, " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.EMPR_CD_NR_EMPRESA FROM " & Session("PREFIXO") & "RELACAO_FINAL INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.ATCA_CD_ATIVIDADE_CARGA"
sql_complemento01=" WHERE " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO="& mega_atual &" AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO=" & proc_atual & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO=" & sub_atual
ssql02=" ORDER BY " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO, " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID.EMPR_CD_NR_EMPRESA"

if len(compl_emp)>0 then
ssql=ssql01+sql_complemento01+ " AND(" + compl_emp + ")" +ssql02
else
ssql="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE MEPR_CD_MEGA_PROCESSO="& mega_atual &" AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual & " ORDER BY ATCA_CD_ATIVIDADE_CARGA, RELA_NR_SEQUENCIA"
end if

'response.write ssql

set rs_ativ=db.execute(ssql)

on error resume next
ativ_atual=rs_ativ("ATCA_CD_ATIVIDADE_CARGA")
trans_atual=rs_ativ("TRAN_CD_TRANSACAO")
seq_atual=rs_ativ("ATCA_NR_SEQUENCIA")
%>
    <%if rs_ativ.eof=false then%>
    <tr> 
      <td width="157" bgcolor="#CCCCCC" align="left" height="20"><font face="Verdana" size="2">ATIVIDADE</font></td>
      <td width="215" bgcolor="#CCCCCC" align="left" height="20"><b><font face="Verdana" size="2">DESCRIÇÃO 
        TRANSA&Ccedil;&Atilde;O </font></b></td>
      <td width="232" bgcolor="#CCCCCC" align="left" height="20"> 
        <p align="left"><b><font face="Verdana" size="2"> TRANSA&Ccedil;&Atilde;O</font></b></p>
      </td>
    </tr>
    <%else
nenhum=1
%>
      <td width="157" height="32"> <font size="2" color="#800000" face="Verdana"> 
        Nenhum registro encontrado </font> </b> 
        <%
end if

ativ_anterior=0

nume_ativ=0

do until rs_ativ.eof=true
nova_ativ=0
%>
    <tr> 
      <%if ativ_anterior <> ativ_atual  then
   nova_ativ=1
	nume_ativ=nume_ativ+1
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA="& ativ_atual)
	PREFIXO=nume_mega & "." & nume_proc & "." & nume_sub & "." & nume_ativ
	ATIVIDADE=rs1("ATCA_TX_DESC_ATIVIDADE")
	%>
      <td width="157" height="21"><font face="Arial" size="1"><%=PREFIXO%> - <%=ATIVIDADE%></font></td>
      <%
   existe=1
   end if%>
      <%IF (TRANS_ANTERIOR<>TRANS_ATUAL) THEN
    set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='"& trim(trans_atual) & "'")
    TRANSACAO = rs2("TRAN_TX_DESC_TRANSACAO")
    if nova_ativ=0 then%>
      <td width="215" height="21"><font face="Arial" size="2">&nbsp;</font></td>
      <%end if%>
      <%IF nova_ativ=1 or TRANS_ANTERIOR<>TRANS_ATUAL THEN%>
      <td width="232" height="21"><font face="Arial" size="1" align="left"> 
        <p align="left"><%=TRANSACAO%>
        </font></td>
      <td width="31" height="21"><font face="Arial" size="1" align="right"> 
        <p align="left"><%=trans_atual%>
        </font></td>
      <%END IF
      END IF
      %>
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

existe=0

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
