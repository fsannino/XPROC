 
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_cenario=request("ID")
str_mega=request("selMegaProcesso")
str_proc=request("selProcesso")
str_sub=request("selSubProcesso")
str_ativ=request("selAtividade")
incl=request("INC")

str_transacao=request("txtTranSelecionada")
str_duplica=request("selDuplicaCenario")

Sub Grava_Transacao(strC, strM, strP, strS, strA, strT,strINCL)
	valor=0
	set rs=db.execute("SELECT MAX(CETR_NR_SEQUENCIA) AS CODIGO FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & strC & "'")
	valor=rs("CODIGO")
	if isnull(valor) then
		valor=1
	ELSE
		valor=valor+1
	end if
	set rs=db.execute("SELECT MAX(CENA_NR_SEQUENCIA_TRANS) AS CODIGO FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & strC & "'")
	valor2=rs("CODIGO")
	if isnull(valor2) then
		valor2 = 10
	ELSE
		valor2 = valor2 + 10
	end if

	set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & strT & "'")

	if not isnull(rs2("TRAN_TX_DESC_TRANSACAO"))then
		desc_transacao=rs2("TRAN_TX_DESC_TRANSACAO")
	else
		desc_transacao=""
	end if	

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "CENARIO_TRANSACAO(CETR_NR_SEQUENCIA, OPES_CD_OPERACAO_ESP, CENA_CD_CENARIO, MEPR_CD_MEGA_PROCESSO, CENA_CD_CENARIO_SEGUINTE, BPPP_CD_BPP, TRAN_CD_TRANSACAO, CENA_NR_SEQUENCIA_TRANS, CETR_TX_DESC_TRANSACAO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, CETR_TX_TIPO_RELACAO) "
	ssql=ssql+"VALUES('" & valor & "',"
	IF strINCL = 3 THEN
       ssql=ssql+" 5,"
	   str_Tp_Relacao = "5"
	ELSE
       ssql=ssql+"NULL,"
	   str_Tp_Relacao = "0"
	END IF   
	ssql=ssql+"'" & trim(ucase(strC)) & "',"
	ssql=ssql+"" & str_mega & ","
	ssql=ssql+"NULL,"
	ssql=ssql+"NULL,"
	ssql=ssql+"'" & strT & "',"
	ssql=ssql+"" & valor2 & ","
	ssql=ssql+"'" & desc_transacao & "',"
	ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE()," & str_Tp_Relacao & ")"

	on error resume next
	db.execute(ssql)
end sub

if len(str_transacao)>1 then

str_valor = str_transacao
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
        
			call Grava_Transacao(str_cenario,str_mega,str_proc,str_sub,str_ativ,str_atual,incl)
			''call grava_log(str_atual,"" & Session("PREFIXO") & "CENARIO_TRANSACAO","I",1)
	   		valor_total=valor_total+1
	   		
        quantos = 0
    End If
    contador = contador + 1
Loop

else

	ssql="INSERT INTO " & Session("PREFIXO") & "CENARIO_TRANSACAO(CETR_NR_SEQUENCIA, OPES_CD_OPERACAO_ESP, CENA_CD_CENARIO, MEPR_CD_MEGA_PROCESSO, CENA_CD_CENARIO_SEGUINTE, BPPP_CD_BPP, TRAN_CD_TRANSACAO, CENA_NR_SEQUENCIA_TRANS, CETR_TX_DESC_TRANSACAO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, CETR_TX_TIPO_RELACAO) "
	ssql=ssql+"(SELECT CETR_NR_SEQUENCIA, OPES_CD_OPERACAO_ESP, '" & ucase(str_cenario) & "', MEPR_CD_MEGA_PROCESSO, CENA_CD_CENARIO_SEGUINTE, BPPP_CD_BPP, TRAN_CD_TRANSACAO, CENA_NR_SEQUENCIA_TRANS, CETR_TX_DESC_TRANSACAO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, CETR_TX_TIPO_RELACAO FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO " 
	ssql=ssql+"WHERE CENA_CD_CENARIO='" & ucase(str_duplica) & "')"
	
	on error resume next
	db.execute(ssql)
	
	''call grava_log(str_cenario,"" & Session("PREFIXO") & "CENARIO_TRANSACAO","I",1)
	
end if
%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body topmargin="0" leftmargin="0">
<form method="POST" action="" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
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
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>

  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Relação 
    Cenário x Transação</font></p>
<p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>

<div align="center">
  <center>
  <table border="0" width="790" height="128">
    <%if err.number=0 then%>
    <tr>
      <td width="233" height="21"><%'=incl%></td>
      <td width="543" height="21" colspan="2"><font face="Verdana" color="#330099" size="2"><b>O
        Registro foi atualizado com Sucesso!</b></font></td>
    </tr>
    <%else%>
    <tr>
      <td width="233" height="21"></td>
      <td width="543" height="21" colspan="2"><b><font face="Verdana" size="2" color="#FF0000">Ocorreu
        um erro na gravação do Registro</font></b></td>
    </tr>
    <%end if%>
    <tr>
      <td width="233" height="21"></td>
      <td width="543" height="21" colspan="2"></td>
    </tr>
    <tr>
            <td width="233" align="right">
              <p align="left"></p>
            </td>
            <td width="48"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0" align="right"></a></td>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
    </tr>
  </center>
  <%IF incl=2 OR incl=3 then%>
  <tr>
            <td width="233" align="right">
              <p align="right"></p>
            </td>
            <td width="48">
              <p align="right"><a href="gerencia_cenario_transa.asp?selMegaProcesso=<%=str_mega%>&selProcesso=<%=str_proc%>&selSubProcesso=<%=str_sub%>&selCenario=<%=str_cenario%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
  <center>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar
              para a tela de Edição de Cenário</font></td>
    </tr>
    </center>
  <%else%>
  <tr>
            <td width="233" align="right">
              <p align="right"></p>
            </td>
            <td width="48">
              <p align="right"><a href="cad_cenario_transacao.asp?option=1&INC=1&ID=<%=str_cenario%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
  <center>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar
              para tela de Relação Cenário x Transação</font></td>
  </tr>
  <%end if%>
    <tr>
      <td width="233"></td>
  </center>
      <td width="48">
  <center>
        <p>&nbsp;</td>
      <td width="489" height="49">
        </td>
    </tr>
  </table>
</div>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>

</html>
