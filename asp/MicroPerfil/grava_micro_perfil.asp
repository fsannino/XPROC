<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("txtAcao") <> "0" then
   str_Acao = request("txtAcao")
else
   str_Acao = ""
end if

if request("selMegaProcesso") <> 0 then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = ""
end if

if request("selFuncPrinc") <> "0" then
   str_FuncPrinc = Ucase(Trim(request("selFuncPrinc")))
else
   str_FuncPrinc = ""
end if

if request("selMacroPerfil") <> 0 then
   str_MacroPerfil = request("selMacroPerfil")
else
   str_MacroPerfil = ""
end if

if request("txtDescM") <> "0" then
   str_Desc = UCase(Trim(request("txtDescM")))
else
   str_Desc = ""
end if

if request("txtdetalM") <> "0" then
   str_DescDet = UCase(Trim(request("txtdetalM")))
else
   str_DescDet = ""
end if

if request("txtespecM") <> "0" then
   str_Espec = UCase(Trim(request("txtespecM")))
else
   str_Espec = ""
end if


str_SQL = ""
str_SQL = str_SQL & " SELECT MAX(MICR_NR_SEQ_MICRO)AS CODIGO "
str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MICRO_PERFIL " 
str_SQL = str_SQL & " where MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso   
set temp=db.execute(str_SQL)
if not isnull(temp("codigo")) then
   sequencia=temp("CODIGO")+1
else
   sequencia=1
end if
temp.close
set temp = Nothing
   
str_SQL = ""
str_SQL = str_SQL & " Select MEPR_TX_ABREVIA from " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL = str_SQL & " where MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
'response.Write(str_SQL)
set rs_Mega = db.execute(str_SQL)
   
ls_AbreviaMega = Trim(rs_Mega("MEPR_TX_ABREVIA"))
ls_CodMicro = ls_AbreviaMega & "." & Right("00000000" & sequencia,5)
ssql="INSERT INTO " & Session("PREFIXO") & "MICRO_PERFIL "
ssql=ssql+ " ( MICR_TX_SEQ_MICRO_PERFIL "
ssql=ssql+ ", MICR_NR_SEQ_MICRO "
ssql=ssql+ ", MICR_TX_DESC_MICRO_PERFIL "
ssql=ssql+ ", MCPR_NR_SEQ_MACRO_PERFIL "
ssql=ssql+ ", MICR_TX_SITUACAO "
ssql=ssql+ ", MEPR_CD_MEGA_PROCESSO "
ssql=ssql+ ", FUNE_CD_FUNCAO_NEGOCIO " 
ssql=ssql+ ", MICR_TX_DESC_DETA_MICRO_PERFIL "
ssql=ssql+ ", MICR_TX_ESPECIFICACAO "
ssql=ssql+ ", ATUA_TX_OPERACAO "
ssql=ssql+ ", ATUA_CD_NR_USUARIO "
ssql=ssql+ ", ATUA_DT_ATUALIZACAO )"
ssql=ssql+ " VALUES( '" & ls_CodMicro & "'"
ssql=ssql+ "," & sequencia 
ssql=ssql+ ",'" & ucase(str_Desc) & "'"
ssql=ssql+ ",'" & ucase(str_MacroPerfil) & "'"
ssql=ssql+ ", 'EE' "
ssql=ssql+ "," & str_MegaProcesso 
ssql=ssql+ ",'" & str_FuncPrinc & "'"
ssql=ssql+ ",'" & str_DescDet & "'"
ssql=ssql+ ",'" & str_Espec & " - " & Session("CdUsuario") & "-" & DATE & "'"
ssql=ssql+ ",'I'"
ssql=ssql+ ",'" & Session("CdUsuario") & "'"
ssql=ssql+ ",GETDATE())"
db.execute(ssql)   
   
SSQL=""
SSQL="SELECT MAX(MHVA_NR_SEQUENCIA_HIST)AS CODIGO FROM " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO WHERE MICR_TX_SEQ_MICRO_PERFIL='" & ls_CodMicro & "'"      		
SET HIST = db.EXECUTE(SSQL)
        		
ATUAL = HIST("CODIGO")
ATUAL = ATUAL + 1
        		
if atual > 1 then
   atual = atual
else
   atual=1
end if
        		
SSQL=""
SSQL="INSERT INTO " & Session("PREFIXO") & "MICRO_HISTORICO_VALIDACAO(MHVA_NR_SEQUENCIA_HIST, MICR_TX_SEQ_MICRO_PERFIL, MHVA_TX_SITUACAO_MICRO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO) "
SSQL=SSQL+"VALUES(" & ATUAL &",'" & ls_CodMicro & "', 'EE', 'I', '" & Session("CdUsuario") & "', GETDATE())"        		
db.EXECUTE(SSQL)
      
db.Close
set db = Nothing

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">


</script>

<script language="javascript" src="../MacroPerfil/js/troca_lista.js"></script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
        <input type="hidden" name="txtImp" size="20"><input type="hidden" name="txtQua" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
      <td colspan="3" height="20">&nbsp;</td>
  </tr>
</table>
        
  
<table width="785" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td width="845">
      </td>
    </tr>
    <tr>
      <td width="845">
        <div align="center"></div>
      </td>
    </tr>
    <tr>
      <td width="845">&nbsp;</td>
    </tr>
  </table>

<table width="102%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp; </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Registro 
      atualizado com sucesso. C&oacute;d Solicita&ccedil;&atilde;o :</b></font><font size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=ls_CodMicro%> </strong></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp; </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td width="6%" height="41"><a href="incluir_micro_perfil.asp?pOPT=1"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td width="94%" height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Inclus&atilde;o de Micro Perfil</font></td>
        </tr>
        <tr> 
          <td height="41"><a href="valida_status.asp?opt=EC&amp;micro=<%=ls_CodMicro%>&amp;acao=C"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font color="#003366" size="2" face="Verdana, Arial, Helvetica, sans-serif">Envia
            para criação no R/3</font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> </td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>

</html>
