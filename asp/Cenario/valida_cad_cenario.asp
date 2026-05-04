<%
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

'response.Write(" AQUI")
'response.Write(Session("Conn_String_Cogest_Gravacao"))
'response.Write("FIM ")

str_mega=request("selMegaProcesso")
str_proc=request("selProcesso")
str_sub=request("selSubProcesso")
str_onda=request("selOnda")
str_classe=request("selClasse")
str_titulo=request("txtTitulo")
str_desc=left(request("txtDescricao"),1000)
str_empresa=request("txtEmpresa")
str_Responsavel=request("txtResp")
str_Dia = request("SelDia")
str_Mes = request("SelMes")
str_Ano = request("SelAno")
'str_DtPrevTermino = str_Mes & "/" & str_Dia & "/" & str_Ano
'IF IsDate(str_DtPrevTermino) = false then
'   response.redirect "envia_msg_tela.asp?opt=0 "
'end if   
''response.write str_DtPrevTermino

str_empresa=right(str_empresa,(len(str_empresa))-1)

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
cod_mega=TRIM(rs("MEPR_TX_ABREVIA"))

set rs_ONDA=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA=" & str_onda)
cod_onda=rs_onda("ONDA_TX_ABREV_ONDA")

cod_proc=right("000" & str_proc, 3)
cod_sub=right("000" & str_sub, 3)
cod_classe=right("000" & str_classe, 3)

contador=0

set rs=db.execute("SELECT MAX(CENA_NR_SEQUENCIA)AS CODIGO FROM " & Session("PREFIXO") & "CENARIO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)

if not isnull(rs("CODIGO")) then
	contador = rs("CODIGO")
end if

if contador=0 then
	contador=1
ELSE
	contador = contador + 1
end if

set rs2=db.execute("SELECT MAX(CENA_NR_SEQUENCIA_ORDEM)AS CODIGO FROM " & Session("PREFIXO") & "CENARIO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND ONDA_CD_ONDA=" & str_onda)

if not isnull(rs2("CODIGO")) then
	contador2 = rs2("CODIGO")
end if

if contador2=0 then
	contador2=1
ELSE
	contador2 = contador2 + 1
end if

cod_sequencia = right("0000" & contador , 4)

codigo = cod_mega & "." & cod_onda & "." & cod_sequencia
'response.write " aqui 2 "

ssql=""
ssql="INSERT INTO " & Session("PREFIXO") & "CENARIO (CENA_TX_TITULO_CENARIO, "
ssql=ssql+" CENA_TX_RESPONSAVEL, "
ssql=ssql+" CENA_TX_DESC_CENARIO, "
ssql=ssql+" PROC_CD_PROCESSO, "
ssql=ssql+" SUPR_CD_SUB_PROCESSO, "
ssql=ssql+" ONDA_CD_ONDA, "
ssql=ssql+" CLCE_CD_NR_CLASSE_CENARIO, "
ssql=ssql+" MEPR_CD_MEGA_PROCESSO, "
ssql=ssql+" CENA_CD_CENARIO, "
ssql=ssql+" ATUA_TX_OPERACAO, "
ssql=ssql+" ATUA_CD_NR_USUARIO, "
ssql=ssql+" ATUA_DT_ATUALIZACAO, "
ssql=ssql+" CENA_DT_DATA_CRIACAO, "
ssql=ssql+" CENA_TX_SITUACAO, "
ssql=ssql+" CENA_NR_SEQUENCIA, "
ssql=ssql+" CENA_NR_SEQUENCIA_ORDEM, "
ssql=ssql+" CENA_TX_SITUACAO_LOTUS, "
ssql=ssql+" CENA_TX_SAIDA, "
ssql=ssql+" CENA_TX_ENTRADA, "
ssql=ssql+" CENA_TX_GERA_FLUXO,"
ssql=ssql+" CENA_TX_EMPRESA_RELAC, "
if request("selAssunto")<>0 then
	ssql=ssql+" CENA_TX_SITUACAO_VALIDACAO, "
	ssql=ssql+" SUMO_NR_CD_SEQUENCIA"
else
	ssql=ssql+" CENA_TX_SITUACAO_VALIDACAO"
end if
ssql=ssql+" ) "
ssql=ssql+"VALUES('" & ucase(str_titulo) & "',"
ssql=ssql+"'" & ucase(str_Responsavel) & "',"
ssql=ssql+"'" & ucase(str_desc) & "',"
ssql=ssql+"" & str_proc & ","
ssql=ssql+"" & str_sub & ","
ssql=ssql+"" & str_onda & ","
ssql=ssql+"" & str_classe & ","
ssql=ssql+"" & str_mega & ","
ssql=ssql+"'" & ucase(codigo) & "',"
ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE(),GETDATE(),"
ssql=ssql+"'EE',"& CONTADOR &","& CONTADOR2 &",'NC','"& request("txtSaida") &"','"& request("txtEntrada") &"' , 'N', '" & str_empresa & "'," 
'ssql=ssql+"'" & UCase(str_Responsavel) & "','" &  str_DtPrevTermino & "',"
if request("selAssunto")<>0 then
	ssql=ssql+" '0' , " & request("selAssunto") & ")"
else
	ssql=ssql+" '0')"
end if
'response.write ssql
'on error resume next
'response.write ssql
db.execute(ssql)

if err.number=0 then
	response.redirect "gerencia_cenario_transa.asp?option=1&INC=1&selCenario=" & codigo & "&selMegaProcesso=" & str_mega & "&selProcesso=" & str_proc & "&selSubProcesso=" & str_sub
else
	response.write err.description
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
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
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
</form>

</body>

</html>