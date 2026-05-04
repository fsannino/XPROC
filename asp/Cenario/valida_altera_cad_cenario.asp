<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

codigo=request("ID")
situacao=request("lotus")
mega=request("selMegaProcesso")
proc=request("selProcesso")
sube=request("selSubProcesso")
onda=request("selOnda")
classe=request("selClasse")
titulo=request("txtTitulo")
desc=request("txtDescricao")
entrada=request("txtEntrada")
saida=request("txtSaida")
empresa=request("txtEmpresa")
str_Responsavel=request("txtResponsavel")
str_Assunto=request("selAssunto")
str_Dia = request("SelDia")
str_Mes = request("SelMes")
str_Ano = request("SelAno")
str_DtPrevTermino = str_Mes & "/" & str_Dia & "/" & str_Ano
IF IsDate(str_DtPrevTermino) = false then
   'response.redirect "envia_msg_tela.asp?opt=0 "
   str_Dia = ""   
end if   

empresa=right(empresa,(len(empresa))-1)

select case situacao
case "NC"
	valor_situacao="NC"
case "CR"
	valor_situacao="RC"
case "RC"
	valor_situacao="RC"
case else
	valor_situacao="NC"
end select

ssql=""
ssql="UPDATE " & Session("PREFIXO") & "CENARIO "
ssql=ssql+"SET MEPR_CD_MEGA_PROCESSO=" & mega & ", "
ssql=ssql+"PROC_CD_PROCESSO=" & proc & ", "
ssql=ssql+"SUPR_CD_SUB_PROCESSO=" & sube & ", "
ssql=ssql+"CLCE_CD_NR_CLASSE_CENARIO=" & classe & ", "

if str_Assunto<>0 then
	ssql=ssql+"SUMO_NR_CD_SEQUENCIA=" & str_Assunto & ", "
end if

ssql=ssql+"CENA_TX_TITULO_CENARIO='" & titulo & "', "
ssql=ssql+"CENA_TX_RESPONSAVEL='" & ucase(request("txtResp")) & "', "
ssql=ssql+"CENA_TX_SITUACAO_LOTUS='" & valor_situacao & "', "
ssql=ssql+"CENA_TX_ENTRADA='" & entrada & "', "
ssql=ssql+"CENA_TX_EMPRESA_RELAC='" & empresa & "', "
if str_Dia <> "" then
   ssql=ssql+"CENA_DT_PREV_TERMINO='" & str_DtPrevTermino & "', "
end if   
ssql=ssql+"CENA_TX_SAIDA='" & saida & "', "
ssql=ssql+"CENA_TX_DESC_CENARIO='" & desc & "' "
ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'"   
ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE()"
ssql=ssql+"WHERE CENA_CD_CENARIO='" & codigo & "'"
'RESPONSE.Write ssql
'on error resume next
db.execute(ssql)

''call grava_log(codigo,"" & Session("PREFIXO") & "CENARIO","A",1)

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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp;</font></p>

  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Alteração
  de Cenário</font></p>

  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<table border="0" width="90%" height="171">
  <%if err.number=0 then%>
  <tr>
    <td width="32%" height="21">&nbsp;</td>
    <td width="70%" colspan="3" height="21"><font face="Verdana" size="2" color="#330099"><b>O
      registro foi alterado com Sucesso</b></font></td>
  </tr>
  <tr>
    <td width="32%" height="21">&nbsp;</td>
    <td width="70%" colspan="3" height="21">&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td width="32%" height="21">&nbsp;</td>
    <td width="70%" colspan="3" height="21"><b><font face="Verdana" size="2" color="#800000">Houve
      um erro na alteração do registro</font></b></td>
  </tr>
  <tr>
    <td width="32%" height="21">&nbsp;</td>
    <td width="70%" colspan="3" height="21">&nbsp;</td>
  </tr>
  <%end if%>
  <tr>
    <td width="32%" height="22">&nbsp;</td>
    <td width="7%" height="22">&nbsp;</td>
    <td width="7%" height="22">
      <p align="right"><font face="Verdana" size="2" color="#330099"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></font></td>
    <td width="56%" height="22"><font face="Verdana" size="2" color="#330099">Volta para a
      tela principal</font></td>
  </tr>
  <tr>
    <td width="32%" height="3">&nbsp;</td>
    <td width="7%" height="3">&nbsp;</td>
    <td width="7%" height="3">&nbsp;</td>
    <td width="56%" height="3">&nbsp;</td>
  </tr>
  <tr>
    <td width="32%" height="1">&nbsp;</td>
    <td width="7%" height="1">&nbsp;</td>
    <td width="7%" height="1">
      <p align="right"><font face="Verdana" size="2" color="#330099"><a href="altera_cenario.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></font></td>
    <td width="56%" height="1"><font face="Verdana" size="2" color="#330099">Volta para a
      tela de alteração de Cenário</font></td>
  </tr>
</table>

  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>

</html>