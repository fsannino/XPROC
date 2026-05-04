 
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_cenario=request("ID")
str_onda=request("selOnda")

set origem=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='"& str_cenario &"'")
str_mega=origem("MEPR_CD_MEGA_PROCESSO")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
cod_mega=TRIM(rs("MEPR_TX_ABREVIA"))

set rs_ONDA=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA=" & str_onda)
cod_onda=rs_onda("ONDA_TX_ABREV_ONDA")

cod_classe=right("000" & str_classe, 3)
contador=0

set rs=db.execute("SELECT MAX(CENA_NR_SEQUENCIA)AS CODIGO FROM " & Session("PREFIXO") & "CENARIO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
if not isnull(rs("CODIGO")) then
	contador = rs("CODIGO")
end if

if contador=0 then
	contador=1
ELSE
	contador=contador+1
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

ssql=""
ssql=ssql+" INSERT INTO " & Session("PREFIXO") & "CENARIO("
ssql=ssql+" CENA_TX_TITULO_CENARIO,CENA_TX_DESC_CENARIO,PROC_CD_PROCESSO"
ssql=ssql+" ,SUPR_CD_SUB_PROCESSO,ONDA_CD_ONDA,CLCE_CD_NR_CLASSE_CENARIO"
ssql=ssql+" ,MEPR_CD_MEGA_PROCESSO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO"
ssql=ssql+" ,ATUA_DT_ATUALIZACAO,CENA_TX_SITUACAO,CENA_NR_SEQUENCIA"
ssql=ssql+" ,CENA_NR_SEQUENCIA_ORDEM,CENA_TX_SITUACAO_LOTUS,CENA_TX_SAIDA"
ssql=ssql+" ,CENA_CD_CENARIO,CENA_TX_ENTRADA,CENA_TX_GERA_FLUXO,CENA_TX_SITUACAO_TESTE"
ssql=ssql+" ,CENA_TX_SITU_DESENHO_TIPO,CENA_TX_SITU_DESENHO_CONF,CENA_TX_SITU_DESENHO_DESE"
ssql=ssql+" ,CENA_TX_CD_CENARIO,CENA_TX_EMPRESA_RELAC,CENA_TX_SITU_DESENHO_TESTE"
ssql=ssql+" ,CENA_TX_SITUACAO_VALIDACAO"
ssql=ssql+" ,CENA_DT_VALIDACAO,CENA_DT_PREV_TERMINO"
ssql=ssql+" ,CENA_TX_RESPONSAVEL,CENA_DT_DATA_CRIACAO,SUMO_NR_CD_SEQUENCIA)"
ssql=ssql+" (SELECT CENA_TX_TITULO_CENARIO,CENA_TX_DESC_CENARIO,PROC_CD_PROCESSO"
ssql=ssql+" ,SUPR_CD_SUB_PROCESSO," & str_onda & ",CLCE_CD_NR_CLASSE_CENARIO"
ssql=ssql+" ,MEPR_CD_MEGA_PROCESSO"
ssql=ssql+" ,'C','" & Session("CdUsuario") & "',GETDATE()"
ssql=ssql+" ,'EE'," & CONTADOR & ", " & CONTADOR2 & ",CENA_TX_SITUACAO_LOTUS"
ssql=ssql+" ,CENA_TX_SAIDA,'" & CODIGO & "',CENA_TX_ENTRADA,CENA_TX_GERA_FLUXO"
ssql=ssql+" ,CENA_TX_SITUACAO_TESTE,CENA_TX_SITU_DESENHO_TIPO,CENA_TX_SITU_DESENHO_CONF"
ssql=ssql+" ,CENA_TX_SITU_DESENHO_DESE,CENA_TX_CD_CENARIO,CENA_TX_EMPRESA_RELAC"
ssql=ssql+" ,CENA_TX_SITU_DESENHO_TESTE,'0' "
ssql=ssql+" ,null,CENA_DT_PREV_TERMINO"
ssql=ssql+" ,CENA_TX_RESPONSAVEL,GETDATE(),SUMO_NR_CD_SEQUENCIA "
ssql=ssql+" FROM " & Session("PREFIXO") & "CENARIO "
ssql=ssql+" WHERE CENA_CD_CENARIO='" & str_cenario & "')"
'response.Write(ssql)
db.execute(ssql)
ssql=""
ssql=ssql+" INSERT INTO " & Session("PREFIXO") & "CENARIO_TRANSACAO(CETR_NR_SEQUENCIA"
ssql=ssql+" , OPES_CD_OPERACAO_ESP, CENA_CD_CENARIO, MEPR_CD_MEGA_PROCESSO"
ssql=ssql+" , CENA_CD_CENARIO_SEGUINTE, BPPP_CD_BPP, TRAN_CD_TRANSACAO"
ssql=ssql+" , CENA_NR_SEQUENCIA_TRANS, CETR_TX_DESC_TRANSACAO, ATUA_TX_OPERACAO"
ssql=ssql+" , ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO, CETR_TX_TIPO_RELACAO) "
ssql=ssql+" (SELECT CETR_NR_SEQUENCIA, OPES_CD_OPERACAO_ESP, '" & codigo & "'"
ssql=ssql+" , MEPR_CD_MEGA_PROCESSO, CENA_CD_CENARIO_SEGUINTE, BPPP_CD_BPP"
ssql=ssql+" , TRAN_CD_TRANSACAO, CENA_NR_SEQUENCIA_TRANS, CETR_TX_DESC_TRANSACAO"
''''ssql=ssql+" , ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO"
ssql=ssql+" ,'C','" & Session("CdUsuario") & "',GETDATE()"
ssql=ssql+" , CETR_TX_TIPO_RELACAO "
ssql=ssql+" FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO " 
ssql=ssql+" WHERE CENA_CD_CENARIO='" & str_cenario & "')"
'response.Write(ssql)
db.execute(ssql)
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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>

<p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Cópia
        de Cenário entre Ondas</font>
</p>
<p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>

<div align="center">
  <center>
  <table border="0" width="790" height="128">
    <%if err.number=0 then%>
    <tr>
      <td width="233" height="21"></td>
      <td width="543" height="21" colspan="2"><b><font face="Verdana" color="#330099" size="2">O
        Registro foi copiado com Sucesso com o código </font><font face="Verdana" color="#330099" size="3"><%=codigo%></font></b></td>
    </tr>
    <%else%>
    <tr>
      <td width="233" height="21"></td>
      <td width="543" height="21" colspan="2"><b><font face="Verdana" size="2" color="#FF0000">Ocorreu
        um erro na cópia do Registro</font></b></td>
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
            <td width="48"><a href="../../indexA.asp"><img src="selecao_F02.gif" width="22" height="20" border="0" align="right"></a></td>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
    </tr>
  </center>
  <tr>
            <td width="233" align="right">
              <p align="right"></p>
            </td>
            <td width="48">
              <p align="right"><a href="copia_cenario_onda.asp"><img src="selecao_F02.gif" width="22" height="20" border="0"></a></td>
  <center>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar
              para a tela Cópia de Cenário entre ondas</font></td>
    </tr>
    </center>
  <tr>
            <td width="233" align="right">
              <p align="right"></p>
            </td>
            <td width="48">
            </td>
  <center>
            <td height="41" width="489"></td>
  </tr>
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




