<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set Conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Atividade = request("selAtividade")
str_Modulo = request("selModulo")
str_Transacao = Request("txtEmpSelecionada")

SSQL1="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & STR_ATIVIDADE & " AND MODU_CD_MODULO=" & str_modulo
set rs_deleta=conn_db.execute(SSQL1)

DO UNTIL RS_DELETA.EOF=TRUE

	TRANS_ATUAL=RS_DELETA("TRAN_CD_TRANSACAO")

	ON ERROR RESUME NEXT
	SSQLD="DELETE FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & STR_ATIVIDADE & " AND MODU_CD_MODULO=" & str_modulo & " AND TRAN_CD_TRANSACAO='" & trans_atual & "'"
	CONN_DB.EXECUTE(SSQLD)
			
	IF ERR.NUMBER=0 THEN
		APAGA=APAGA+1
	END IF
	
	GERAL = GERAL + 1
	
	RS_DELETA.MOVENEXT
	
LOOP

IF LEN(APAGA)=0 THEN
	APAGA=0
END IF

Sub Grava_Nova_Atividade(strT,strA,strM)
	
	str_SQL_Nova_Sub_Empr = ""
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & "INSERT INTO " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA VALUES('"
   	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & strT & "'," & strA & ", " & strM & ", "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
   	
   	conn_db.Execute(str_SQL_Nova_Sub_Empr)
	
	strChave = CStr(strT) & " " & CStr(strA) & " " & CStr(strM) 
	'call grava_log(strChave,"MODU_ATIV_TRA_CARGA","I",0)

end sub

'guarda o conteúdo da String
str_valor = str_Transacao

'Coloca uma virgula no fim de string, se não houver
if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if

'Pega o tamanho da string
tamanho = Len(str_valor)

'Retira a vírgula do início da string, se houver
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If

'Atualiza o tamanho da string
tamanho = Len(str_valor)

'Inicializa o Contador
contador = 1

'início da Rotina

Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)

    If str_temp = "," Then
    
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        
	   'Aqui entra o que vc quer fazer com o caracter em questão!
	
		ssql_existe="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & STR_ATIVIDADE & " AND MODU_CD_MODULO=" & str_modulo & " AND TRAN_CD_TRANSACAO='" & str_atual& "'"
		
		set existe=conn_db.execute(ssql_existe)
		
		quantosR = quantosR + 1
		
		if existe.eof=true then		   

			call Grava_Nova_Atividade(str_atual,str_Atividade,str_Modulo)
	   		valor_total=valor_total+1
	   		
	   	end if

        quantos = 0
    End If
    contador = contador + 1
Loop

if len(valor_total)=0 then
	valor_total=0
end if

if tem_erro<>0 then
	valor_total=0
end if

SSQL2="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & STR_ATIVIDADE & " AND MODU_CD_MODULO=" & str_modulo
set rs_atual=conn_db.execute(SSQL1)

do until rs_atual.eof=true
	atual=atual+1
	rs_atual.movenext
loop

if quantosR<>atual then
	tem_erro=1
end if

conn_db.Close
set conn_db = Nothing
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
    <td width="36%" valign="top" height="65"> 
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
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27">&nbsp;</td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="78%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp; </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%"></td>
    <td width="70%">
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de Registros Anterior : <%=geral%></font>
    </td>
    <td width="16%"></td>
  </tr>
  <tr> 
    <td width="14%"></td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total
      de Registros Atual : <%=atual%></font></td>
    <td width="16%"></td>
  </tr>
  <tr> 
    <td width="14%"></td>
    <td width="70%"></td>
    <td width="16%"></td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%if tem_erro=0 then%>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Relação
    entre Agrupamento ( Master List R/3 ) x Atividade x Transação&nbsp; realizada com sucesso!</b></font> 
      
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%else%>
  <tr> 
    <td width="14%"></td>
    <td width="70%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Alguns
      registros não puderam ser atualizados devido à violações de
      relacionamento.&nbsp;</font></b></td>
    <td width="16%"></td>
  </tr>
  <%end if%>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="relacao_master_.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">&nbsp;Volta 
            para tela de Rela&ccedil;&atilde;o Master List R/3 / Atividade /
            Transação</font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">&nbsp;Volta 
            para tela Principal</font></td>
        </tr>
      </table>
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set Conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Atividade = request("selAtividade")
str_Modulo = request("selModulo")
str_Transacao = Request("txtEmpSelecionada")

SSQL1="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & STR_ATIVIDADE & " AND MODU_CD_MODULO=" & str_modulo
set rs_deleta=conn_db.execute(SSQL1)

DO UNTIL RS_DELETA.EOF=TRUE

	TRANS_ATUAL=RS_DELETA("TRAN_CD_TRANSACAO")

	ON ERROR RESUME NEXT
	SSQLD="DELETE FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & STR_ATIVIDADE & " AND MODU_CD_MODULO=" & str_modulo & " AND TRAN_CD_TRANSACAO='" & trans_atual & "'"
	CONN_DB.EXECUTE(SSQLD)
			
	IF ERR.NUMBER=0 THEN
		APAGA=APAGA+1
	END IF
	
	GERAL = GERAL + 1
	
	RS_DELETA.MOVENEXT
	
LOOP

IF LEN(APAGA)=0 THEN
	APAGA=0
END IF

Sub Grava_Nova_Atividade(strT,strA,strM)
	
	str_SQL_Nova_Sub_Empr = ""
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & "INSERT INTO " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA VALUES('"
   	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & strT & "'," & strA & ", " & strM & ", "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
   	
   	conn_db.Execute(str_SQL_Nova_Sub_Empr)
	
	strChave = CStr(strT) & " " & CStr(strA) & " " & CStr(strM) 
	'call grava_log(strChave,"MODU_ATIV_TRA_CARGA","I",0)

end sub

'guarda o conteúdo da String
str_valor = str_Transacao

'Coloca uma virgula no fim de string, se não houver
if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if

'Pega o tamanho da string
tamanho = Len(str_valor)

'Retira a vírgula do início da string, se houver
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If

'Atualiza o tamanho da string
tamanho = Len(str_valor)

'Inicializa o Contador
contador = 1

'início da Rotina

Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)

    If str_temp = "," Then
    
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        
	   'Aqui entra o que vc quer fazer com o caracter em questão!
	
		ssql_existe="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & STR_ATIVIDADE & " AND MODU_CD_MODULO=" & str_modulo & " AND TRAN_CD_TRANSACAO='" & str_atual& "'"
		
		set existe=conn_db.execute(ssql_existe)
		
		quantosR = quantosR + 1
		
		if existe.eof=true then		   

			call Grava_Nova_Atividade(str_atual,str_Atividade,str_Modulo)
	   		valor_total=valor_total+1
	   		
	   	end if

        quantos = 0
    End If
    contador = contador + 1
Loop

if len(valor_total)=0 then
	valor_total=0
end if

if tem_erro<>0 then
	valor_total=0
end if

SSQL2="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & STR_ATIVIDADE & " AND MODU_CD_MODULO=" & str_modulo
set rs_atual=conn_db.execute(SSQL1)

do until rs_atual.eof=true
	atual=atual+1
	rs_atual.movenext
loop

if quantosR<>atual then
	tem_erro=1
end if

conn_db.Close
set conn_db = Nothing
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
    <td width="36%" valign="top" height="65"> 
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
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27">&nbsp;</td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="78%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp; </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%"></td>
    <td width="70%">
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de Registros Anterior : <%=geral%></font>
    </td>
    <td width="16%"></td>
  </tr>
  <tr> 
    <td width="14%"></td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total
      de Registros Atual : <%=atual%></font></td>
    <td width="16%"></td>
  </tr>
  <tr> 
    <td width="14%"></td>
    <td width="70%"></td>
    <td width="16%"></td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%if tem_erro=0 then%>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Relação
    entre Agrupamento ( Master List R/3 ) x Atividade x Transação&nbsp; realizada com sucesso!</b></font> 
      
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <%else%>
  <tr> 
    <td width="14%"></td>
    <td width="70%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Alguns
      registros não puderam ser atualizados devido à violações de
      relacionamento.&nbsp;</font></b></td>
    <td width="16%"></td>
  </tr>
  <%end if%>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="relacao_master_.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">&nbsp;Volta 
            para tela de Rela&ccedil;&atilde;o Master List R/3 / Atividade /
            Transação</font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">&nbsp;Volta 
            para tela Principal</font></td>
        </tr>
      </table>
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
