<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<%
str_mod = Request("selModulo")
str_ativ = Request("selAtividade")
str_tran = Request("selTransacao")
str_emp = Request("selEmpresa")

if str_mod<>0 then
	valor="Master List R/3"
	valor_sim="Agrupamento ( Master List R/3 ) Excluído com Sucesso!"
	ssql="DELETE FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & str_mod
	TABf="" & Session("PREFIXO") & "MODULO_R3"
	REGf=str_mod
end if

if str_ativ<>0 then
	valor="Atividade"
	valor_sim="Atividade Excluída com Sucesso!"
	ssql="DELETE FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & str_ativ
	TABf="" & Session("PREFIXO") & "ATIVIDADE_CARGA"
	REGf=str_ativ
end if

if len(str_tran)<>0 then
	valor="Transação"
	valor_sim="Transação Excluída com Sucesso!"
	ssql="DELETE FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & str_tran & "'"
	TABf="" & Session("PREFIXO") & "TRANSACAO"
	REGf=str_tran
end if

if str_emp<>0 then
	valor="Empresa"
	valor_sim="Empresa Excluída com Sucesso!"
	ssql="DELETE FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE WHERE EMPR_CD_NR_EMPRESA=" & str_emp
	TABf="" & Session("PREFIXO") & "EMPRESA_UNIDADE"
	REGf=str_empr
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

on error resume next
'response.write ssql
db.execute(ssql)
'call grava_log(REGf,TABf,"D",1)

if err.number<>0 then
	valor_erro="Não foi possível excluir o registro, pois o mesmo possui registros relacionados à ele"
else
	valor_ok=valor_sim
end if
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" action="valida_exclusao.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
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
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="2%"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:Confirma()"></td>
      <td height="20" width="43%"><font color="#330099" size="2" face="Verdana"><b>Excluir</b></font></td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="40%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="17%">&nbsp;</td>
      <td width="69%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="17%">&nbsp;</td>
      <td width="69%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Exclusão
        de <%=valor%></font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="851" border="0" cellspacing="0" cellpadding="0" height="131">
    <tr> 
      <td width="192" height="21">&nbsp;</td>
      <td width="440" colspan="3" height="21">&nbsp;</td>
      <td width="65" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="192" height="21"></td>
      <td width="440" colspan="3" height="21"><b><font size="2" face="Verdana" color="#000099"><%=valor_ok%></font></b></td>
      <td width="65" height="21"></td>
    </tr>
    <tr> 
      <td width="192" height="21"></td>
      <td width="440" colspan="3" height="21"><b><font size="2" face="Verdana" color="#FF0000"><%=valor_erro%></font></b></td>
      <td width="65" height="21"></td>
    </tr>
    <tr> 
      <td width="192" height="21"></td>
      <td width="440" colspan="3" height="21"></td>
      <%select case str_id
      case 1%>
      <%case 2%>
      <td width="65" height="21"></td>
      <%case 3%>
      <%case 4%><%end select%>
    </tr>
    <tr>
            <td height="26" width="192"></td>
            <td height="26" width="12">
              <p align="right"></p>
      </td>
            <td height="26" width="17">
              <a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a>
      </td>
            <td height="26" width="427">
              <p align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">&nbsp;
              Volta 
              para tela Principal</font></p>
            </td>
    </tr>
    <tr> 
      <td width="192" height="21">&nbsp;</td>
      <td width="440" colspan="3" height="21">&nbsp;</td>
      <td width="65" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="192" height="21">&nbsp;</td>
      <td width="440" colspan="3" height="21">&nbsp;</td>
      <td width="65" height="21">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<%
str_mod = Request("selModulo")
str_ativ = Request("selAtividade")
str_tran = Request("selTransacao")
str_emp = Request("selEmpresa")

if str_mod<>0 then
	valor="Master List R/3"
	valor_sim="Agrupamento ( Master List R/3 ) Excluído com Sucesso!"
	ssql="DELETE FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & str_mod
	TABf="" & Session("PREFIXO") & "MODULO_R3"
	REGf=str_mod
end if

if str_ativ<>0 then
	valor="Atividade"
	valor_sim="Atividade Excluída com Sucesso!"
	ssql="DELETE FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & str_ativ
	TABf="" & Session("PREFIXO") & "ATIVIDADE_CARGA"
	REGf=str_ativ
end if

if len(str_tran)<>0 then
	valor="Transação"
	valor_sim="Transação Excluída com Sucesso!"
	ssql="DELETE FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & str_tran & "'"
	TABf="" & Session("PREFIXO") & "TRANSACAO"
	REGf=str_tran
end if

if str_emp<>0 then
	valor="Empresa"
	valor_sim="Empresa Excluída com Sucesso!"
	ssql="DELETE FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE WHERE EMPR_CD_NR_EMPRESA=" & str_emp
	TABf="" & Session("PREFIXO") & "EMPRESA_UNIDADE"
	REGf=str_empr
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

on error resume next
'response.write ssql
db.execute(ssql)
'call grava_log(REGf,TABf,"D",1)

if err.number<>0 then
	valor_erro="Não foi possível excluir o registro, pois o mesmo possui registros relacionados à ele"
else
	valor_ok=valor_sim
end if
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" action="valida_exclusao.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
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
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="2%"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:Confirma()"></td>
      <td height="20" width="43%"><font color="#330099" size="2" face="Verdana"><b>Excluir</b></font></td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="40%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="17%">&nbsp;</td>
      <td width="69%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="17%">&nbsp;</td>
      <td width="69%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Exclusão
        de <%=valor%></font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="851" border="0" cellspacing="0" cellpadding="0" height="131">
    <tr> 
      <td width="192" height="21">&nbsp;</td>
      <td width="440" colspan="3" height="21">&nbsp;</td>
      <td width="65" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="192" height="21"></td>
      <td width="440" colspan="3" height="21"><b><font size="2" face="Verdana" color="#000099"><%=valor_ok%></font></b></td>
      <td width="65" height="21"></td>
    </tr>
    <tr> 
      <td width="192" height="21"></td>
      <td width="440" colspan="3" height="21"><b><font size="2" face="Verdana" color="#FF0000"><%=valor_erro%></font></b></td>
      <td width="65" height="21"></td>
    </tr>
    <tr> 
      <td width="192" height="21"></td>
      <td width="440" colspan="3" height="21"></td>
      <%select case str_id
      case 1%>
      <%case 2%>
      <td width="65" height="21"></td>
      <%case 3%>
      <%case 4%><%end select%>
    </tr>
    <tr>
            <td height="26" width="192"></td>
            <td height="26" width="12">
              <p align="right"></p>
      </td>
            <td height="26" width="17">
              <a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a>
      </td>
            <td height="26" width="427">
              <p align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">&nbsp;
              Volta 
              para tela Principal</font></p>
            </td>
    </tr>
    <tr> 
      <td width="192" height="21">&nbsp;</td>
      <td width="440" colspan="3" height="21">&nbsp;</td>
      <td width="65" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="192" height="21">&nbsp;</td>
      <td width="440" colspan="3" height="21">&nbsp;</td>
      <td width="65" height="21">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
