<%
mega=request("selMegaProcesso")
titulo_funcao=request("txtfuncao")
descricao_funcao=request("txtdescfuncao")
generica=request("selGenerica")
str_EmUso = request("chkEmUso")
sub_modulo=request("txtAss")
strAntec = request("chkAntec")

IF strAntec <>1 THEN
	strAntec = 0
END IF

if request("selFuncaoPai") <> "0" then
   pai=request("selFuncaoPai")
else
   pai= "0"
end if   
'on error resume next
'if pai<>0 then
'if err.number<>0 then
if pai <> "0" then
	valor_pai="'" & pai & "'"
	valor_ref=1
'end if
else
	valor_pai=""
end if

str_quali=request("txtqua")
str_imp=request("txtImp")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

atual=0

set rs=db.execute("SELECT MAX(FUNE_NR_SEQUENCIA)AS CODIGO FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)

if not isnull(rs("CODIGO")) then
	ATUAL = rs("CODIGO")
end if

if atual=0 then
	atual=1
else
	atual=atual+1
end if

codigo=right("00"& atual,3)

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)

if generica=1 then
	codigo=ucase(rs("MEPR_TX_ABREVIA"))&"."& codigo
else
	codigo=ucase(rs("MEPR_TX_ABREVIA"))&"."& codigo
end if

if generica=1 then
	valor_generica="G"
else
	valor_generica="N"
end if

if str_EmUso = 1 then
	valor_EmUso= 1
else
	valor_EmUso= 0
end if

if valor_pai="" then
	valor_pai="'" & codigo & "'"
	valor_ref=0
END IF

ssql=""
ssql="INSERT INTO " & Session("PREFIXO") & "FUNCAO_NEGOCIO ("
ssql=ssql & " FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNE_NM_ANTECIPADA, "
ssql=ssql & " FUNE_TX_DESC_FUNCAO_NEGOCIO, "
ssql=ssql & " MEPR_CD_MEGA_PROCESSO, ATUA_TX_OPERACAO, "
ssql=ssql & " ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, "
ssql=ssql & " FUNE_CD_FUNCAO_NEGOCIO, FUNE_TX_TP_FUN_NEG, "
ssql=ssql & " FUNE_NR_SEQUENCIA, FUNE_IN_BLOQUEADO, FUNE_CD_FUNCAO_NEGOCIO_PAI, FUNE_TX_INDICA_REFERENCIADA, FUNE_TX_INDICA_EM_USO, FUNE_TX_LIBERADO_MANUT_PERF, FUNE_TX_TIPO_CLASS"
ssql=ssql & ") VALUES ('" & ucase(titulo_funcao) & "', " & strAntec & ", "
ssql=ssql+"'" & ucase(descricao_funcao) & "',"
ssql=ssql+"" & mega & ","
ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE(),"
ssql=ssql+"'" & codigo & "','" & valor_generica & "'," & atual & ",'N', " & valor_pai & ",'" & valor_ref & "','" & valor_EmUso  & "','0','0')"

'response.write ssql

db.execute(ssql)

Sub Grava_quali(strF, strQ)

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "FUN_NEG_TP_QUA "
	ssql=ssql+"VALUES(" & strQ & ","
	ssql=ssql+"'" & ucase(strF) & "',"
	ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE())"

	db.execute(ssql)
	
end sub

Sub Grava_Imp(strF, strI)

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU "
	ssql=ssql+"VALUES('" & ucase(strF) & "',"
	ssql=ssql+"'" & strI & "',"
	ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"
    'response.write ssql
	db.execute(ssql)
	
end sub

Sub Grava_Mod(strF, strI)

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "FUNCAO_NEGOCIO_SUB_MODULO "
	ssql=ssql+"VALUES('" & ucase(strF) & "',"
	ssql=ssql+"'" & strI & "',"
	ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"
    
	db.execute(ssql)
	
end sub

str_valor = str_quali

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
        
			call Grava_quali(codigo,str_atual)
	   		valor_total=valor_total+1
	   		
        quantos = 0
    End If
    contador = contador + 1
Loop

if str_imp <> "" then

   str_valor = str_imp

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
        
	     call Grava_Imp(codigo,str_atual)
	     valor_total=valor_total+1
	   		
         quantos = 0
      End If
      contador = contador + 1
   Loop
else
   str_SQL = ""
   str_SQL = str_SQL & " Select "
   str_SQL = str_SQL & " AGLU_CD_AGLUTINADO "
   str_SQL = str_SQL & " , AGLU_SG_AGLUTINADO "
   str_SQL = str_SQL & " FROM dbo.ORGAO_AGLUTINADOR "
   set rds_Org_Aglu = db.execute(str_SQL)
   'RESPONSE.Write(str_SQL)
   do while not rds_Org_Aglu.EOF
      call Grava_Imp(codigo,rds_Org_Aglu("AGLU_CD_AGLUTINADO"))
      rds_Org_Aglu.movenext
   Loop
   
end if

str_valor = sub_modulo

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
        
			call Grava_Mod(codigo,str_atual)
	   		valor_total=valor_total+1
	   		
        quantos = 0
    End If
    contador = contador + 1
Loop

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body topmargin="0" leftmargin="0">
<form method="POST" action="" name="frm1">
<input type="hidden" name="txtpub" size="20"><input type="hidden" name="txtQua" size="20">
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
          <td width="26"></td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Cadastro
        de Fun&ccedil;&atilde;o R/3</font></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2">O
Registro foi incluído com sucesso com o </font><font face="Verdana" color="#330099" size="2"> Código
</font><font face="Verdana" color="#330099" size="3"> <%=codigo%></font></b></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="889" height="119">
  <tr>
    <td width="287" height="37"></td>
            <td width="26" height="37" align="left" valign="middle"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
            <td width="556" height="37" align="left" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
  </tr>
  <tr>
    <td width="287" height="37"></td>
            <td width="26" height="37" align="left" valign="middle">
      <p align="right"><a href="cad_funcao.asp"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
            <td width="556" height="37" align="left" valign="middle"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar
              para a tela de Cadastro de Fun&ccedil;&atilde;o R/3</font></td>
  </tr>
  <tr>
    <td width="287" height="37"></td>
    <%if request("pai")=0 then%>
    <td width="26" height="37" align="left" valign="middle">
      <p align="right"><a href="cad_funcao_transacao2.asp?selMegaProcesso=<%=mega%>&selFuncao=<%=codigo%>"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
    <td width="556" height="37" align="left" valign="middle">
      <p style="margin-top: 0; margin-bottom: 0"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Relacionar
      Fun&ccedil;&atilde;o R/3 x Transação</font></p>
      <p style="margin-top: 0; margin-bottom: 0"></td>
    <%end if%>
  </tr>
</table>
</form>

<p>&nbsp;</p>

</body>

</html>