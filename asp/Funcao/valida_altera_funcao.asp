<%
funcao=request("funcao")
mega=request("selMegaProcesso")
titulo_funcao=request("txtfuncao")
descricao_funcao=request("txtdescfuncao")
generica=request("selGenerica")
str_EmUso = request("chkEmUso")
'response.Write(str_EmUso)
sub_modulo=request("txtAss")
sub_Obs_Especifica = request("txtObs_Especifica")
'pai=request("selFuncaoPai")

strAntec = request("chkAntec")

IF strAntec <>1 THEN
	strAntec = 0
END IF

if request("selFuncaoPai") <> "0" then
   pai=request("selFuncaoPai")
else
   pai= "0"
end if   
if pai <> "0" then
	valor_pai="'" & pai & "'"
	valor_ref=1
else
	valor_pai=""
end if

'on error resume next
'if pai<>0 then
'if err.number<>0 then
'	valor_pai="'" & pai & "'"
'	valor_ref=1
'end if
'else
'	valor_pai=""
'end if

if valor_pai="" then
	testa_pai = funcao
	valor_pai="'" & funcao & "'"
	valor_ref=0
END IF

str_quali=request("txtqua")
str_Imp=request("txtImp")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_TP_QUA WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
db.execute("DELETE FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO_SUB_MODULO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")

if generica=1 then
	valor_generica="G"
else
   if generica=0 then
      valor_generica="G"
   else
	  valor_generica="N"
   end if
end if

if generica="" then
   valor_generica="N"
end if

if str_EmUso = 1 then
   valor_EmUso = 1   
else
   if str_EmUso = 0 then
      valor_EmUso = 1   
   else
      valor_EmUso = 0
   end if
end if
if str_EmUso="" then
   valor_EmUso= 0
end if

'response.write " valor gene "
'response.write valor_generica

codigo=funcao

valor_filho="'" & trim(codigo) & "'"

if trim(valor_filho)<>trim(valor_pai) then
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & codigo & "'")
	db.execute("UPDATE " & Session("PREFIXO") & "MACRO_PERFIL SET MCPE_TX_SITUACAO='MR' WHERE FUNE_CD_FUNCAO_NEGOCIO='" & codigo & "'")	
end if

ssql=""
ssql="UPDATE " & Session("PREFIXO") & "FUNCAO_NEGOCIO "
ssql=ssql & "SET FUNE_TX_TITULO_FUNCAO_NEGOCIO='" & ucase(titulo_funcao) & "', "
ssql=ssql+"FUNE_TX_DESC_FUNCAO_NEGOCIO='" & ucase(descricao_funcao) & "',"
ssql=ssql+"FUNE_CD_FUNCAO_NEGOCIO_PAI=" & valor_pai & ","
ssql=ssql+"FUNE_NM_ANTECIPADA=" & strAntec  & ","
ssql=ssql+"FUNE_TX_TP_FUN_NEG='" & valor_generica & "', "
ssql=ssql+"MEPR_CD_MEGA_PROCESSO=" & mega & ", "
ssql=ssql+"FUNE_TX_INDICA_REFERENCIADA='" & valor_ref & "', "
ssql=ssql+"FUNE_TX_INDICA_EM_USO='" & valor_EmUso & "' "
ssql=ssql+" WHERE FUNE_CD_FUNCAO_NEGOCIO='" & codigo & "'"

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
	ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE())"
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
			''call grava_log(str_atual,"" & Session("PREFIXO") & "FUN_NEG_TP_QUA","I",1)

	   		valor_total=valor_total+1
	   		
        quantos = 0
    End If
    contador = contador + 1
Loop

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
		''call grava_log(str_atual,"" & Session("PREFIXO") & "FUN_NEG_ABR_IM","I",1)
	   	valor_total=valor_total+1	   		
        quantos = 0
    End If
    contador = contador + 1
Loop

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
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Alteração
        de Fun&ccedil;&atilde;o R/3</font></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2">O
Registro foi atualizado com sucesso!</font></b></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="844" height="155">
  <tr>
    <td width="350" height="37"></td>
            <td width="28" height="37"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
            <td height="37" width="446"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
  </tr>
  <tr>
    <td width="350" height="37"></td>
            <td width="28" height="37">
              <p align="right"><a href="seleciona_funcao.asp?pOPT=1"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
            <td height="37" width="446"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar
              para a tela de Alteração de Fun&ccedil;&atilde;o R/3</font></td>
  </tr>
  <%
  if trim(funcao)=trim(testa_pai) then
  %>
  <tr>
    <td width="350" height="44"></td>
    <td width="28" height="44">
      <p align="right"><a href="cad_funcao_transacao2.asp?selMegaProcesso=<%=mega%>&selFuncao=<%=codigo%>"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
    <td width="446" height="44">
      <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
      <p style="margin-top: 0; margin-bottom: 0"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Relacionar
      Fun&ccedil;&atilde;o R/3 x Transação</font></p>
      <p style="margin-top: 0; margin-bottom: 0"></td>
  </tr>
  <%
  else
  db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
  %>
  <%end if%>
  <tr>
    <td width="350" height="21"></td>
    <td width="28" height="21"></td>
    <td width="446" height="21"></td>
  </tr>
</table>
  </form>

<p>&nbsp;</p>

</body>

</html>